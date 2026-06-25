"""
reformat_document.py

Applies the full document reformatting and correction prompt to a .docx file.

Strategy:
1. Send the docx to Claude as a base64 document along with the full reformatting prompt
2. Claude analyses the document and returns a structured JSON correction plan
3. This script applies every correction deterministically using python-docx + lxml
4. The corrected document is saved to the output path

The reformatting prompt is embedded in this script so it is the single source of truth.
"""

import argparse
import base64
import json
import os
import re
import shutil
import sys
import time
import zipfile
import tempfile
from copy import deepcopy

import anthropic
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree


# ── Constants ─────────────────────────────────────────────────────────────────

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Style names as used in this document suite
STYLE_H1        = "Heading1Numbered"
STYLE_H2        = "Heading2Numbered"
STYLE_NORMAL    = "Normal"
STYLE_BODY      = "NormalNumberedtext"
STYLE_BULLET_H1 = "NumberedBulletL1"
STYLE_BULLET_H2 = "NumberedBulletL3"
STYLE_TABLE     = "TableContents"

BLUE_COLOUR = "0059D1"
EM_DASH     = "\u2014"
EN_DASH     = "\u2013"

# Table widths in twips (dxa)
TABLE_WIDTH_H1  = 9777   # under level 1 heading
TABLE_WIDTH_H2  = 9026   # under level 2 heading
TABLE_INDENT_H1 = 0
TABLE_INDENT_H2 = 704

# Cell margins in twips
CELL_TOP    = 36
CELL_BOTTOM = 36
CELL_LEFT   = 108
CELL_RIGHT  = 108

REFORMATTING_PROMPT = """You are analysing a Word document (.docx) to produce a correction plan.
Read the entire document content carefully and return a JSON correction plan following the rules below.
Do not summarise or describe — return only valid JSON.

SCOPE:
- Apply corrections only to body content after the Table of Contents
- Do NOT change: cover page (page 1), document info page (page 2), TOC page, headers, footers

Return a JSON object with this structure:

{
  "paragraphs": [
    {
      "index": <integer — 0-based paragraph index in the document body>,
      "action": "<keep|restyle|remove|blank|fix_text>",
      "new_style": "<style name to apply, or null>",
      "new_text": "<corrected text if action is fix_text, or null>",
      "notes": "<brief reason>"
    }
  ],
  "tables": [
    {
      "index": <integer — 0-based table index>,
      "width_twips": <9777 for under H1, 9026 for under H2>,
      "indent_twips": <0 for under H1, 704 for under H2>,
      "context": "<under_h1 or under_h2>"
    }
  ],
  "text_fixes": [
    {
      "find": "<exact text to find>",
      "replace": "<replacement text>",
      "scope": "<body|tables|all>"
    }
  ],
  "traceability_rows": [
    {
      "row_index": <integer — 0-based row index in the traceability table, skip header row>,
      "section": "<section number(s) where requirement is addressed, or blank>",
      "status": "<Full|Partial|Not met>"
    }
  ]
}

RULES FOR PARAGRAPHS:

1. HEADING LEVEL 1 (Heading1Numbered):
   - Strip any leading number + whitespace from heading text (e.g. "1  Purpose" → "Purpose")
   - One blank NormalNumberedtext paragraph must precede it (except very first body element)
   - action: "restyle" with new_style "Heading1Numbered" and fix_text to strip the number

2. HEADING LEVEL 2 (Heading2Numbered):
   - Strip any leading number.number + whitespace (e.g. "1.1  Purpose" → "Purpose")
   - One blank NormalNumberedtext paragraph must precede it
   - action: "restyle" with new_style "Heading2Numbered"

3. BODY PARAGRAPHS UNDER H1 (before first H2 in that section):
   - Use style "Normal", size 10
   - action: "restyle" with new_style "Normal"

4. BODY PARAGRAPHS UNDER H2:
   - Use style "NormalNumberedtext", size 10
   - action: "restyle" with new_style "NormalNumberedtext"

5. BULLETS UNDER H2:
   - Use style "NumberedBulletL3"
   - action: "restyle" with new_style "NumberedBulletL3"

6. BULLETS UNDER H1:
   - Use style "NumberedBulletL1"
   - action: "restyle" with new_style "NumberedBulletL1"

7. BLANK SPACING PARAGRAPHS:
   - Must use style "NormalNumberedtext"
   - Never use "Default" or "Normal" for blank spacers
   - action: "restyle" with new_style "NormalNumberedtext"

8. CONSECUTIVE BLANK PARAGRAPHS:
   - Remove the duplicate — keep only one
   - action: "remove" for the duplicate

9. API ERROR TEXT (Python dict strings like {'text': '...', 'suggested': True}):
   - Extract only the value of the 'text' key as plain text
   - action: "fix_text" with new_text set to the extracted text value

RULES FOR TEXT FIXES:

1. EM-DASHES (— U+2014) and EN-DASHES (– U+2013):
   - Replace with plain hyphen (-)
   - Add as text_fixes entries with scope "all"

2. DOCUMENT NUMBER CODES (RTPR-HSE-XXX-NNNN pattern):
   - Replace with: [INSERT: Project Document Number]
   - Add as text_fixes entries

3. ROUND BRACKETS AROUND PLACEHOLDERS ([INSERT: ...]):
   - If any placeholder appears as ([INSERT: ...]), remove the surrounding ()
   - Add as text_fixes entries

RULES FOR TABLES:
For each table, determine if it sits under a H1 or H2 heading and set widths accordingly.
Under H1: width=9777, indent=0
Under H2: width=9026, indent=704

RULES FOR TRACEABILITY TABLE:
Find the "Source Requirements Traceability" section near the end.
For each data row (skip header), assess:
- Column 1 (Section): identify the section number where the requirement is addressed
- Column 4 (Status): Full / Partial / Not met

Return only valid JSON. No markdown, no preamble."""


# ── Claude API call ───────────────────────────────────────────────────────────

def get_correction_plan(client, docx_path):
    """Send the document to Claude and get back a correction plan JSON."""

    with open(docx_path, "rb") as f:
        docx_bytes = f.read()
    docx_b64 = base64.standard_b64encode(docx_bytes).decode("utf-8")

    # Also extract plain text for Claude to read the content
    doc = Document(docx_path)
    plain_text_parts = []
    for i, para in enumerate(doc.paragraphs):
        style = para.style.name if para.style else "Unknown"
        plain_text_parts.append(f"[P{i}|{style}] {para.text}")
    for i, table in enumerate(doc.tables):
        plain_text_parts.append(f"[TABLE {i}]")
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            plain_text_parts.append(f"  {row_text}")

    plain_text = "\n".join(plain_text_parts)

    # Truncate if very long
    if len(plain_text) > 30000:
        plain_text = plain_text[:30000] + "\n[... truncated ...]"

    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=8000,
                system="You are a document correction specialist. Output only valid JSON.",
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                "data": docx_b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": (
                                f"DOCUMENT PARAGRAPH INDEX (for correction plan):\n{plain_text}\n\n"
                                f"{REFORMATTING_PROMPT}"
                            )
                        }
                    ]
                }]
            )

            raw = response.content[0].text.strip()
            raw = re.sub(r"^```(?:json)?\s*", "", raw)
            raw = re.sub(r"\s*```$", "", raw)
            return json.loads(raw)

        except json.JSONDecodeError as e:
            print(f"  JSON error attempt {attempt+1}: {e}")
            if attempt < 2:
                time.sleep(5)
        except Exception as e:
            print(f"  API error attempt {attempt+1}: {e}")
            if attempt < 2:
                time.sleep(10)

    return None


# ── XML helpers ───────────────────────────────────────────────────────────────

def set_run_font_arial(r_elem):
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_elem.insert(0, rPr)
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rPr.insert(0, fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), "Arial")


def set_para_style(p_elem, style_id):
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), style_id)


def get_para_style(p_elem):
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return None
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return None
    return pStyle.get(qn("w:val"))


def get_para_text(p_elem):
    return "".join(t.text or "" for t in p_elem.iter(qn("w:t")))


def set_run_size(r_elem, half_pts):
    """Set sz and szCs on a run. half_pts=20 means 10pt, half_pts=16 means 8pt."""
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_elem.insert(0, rPr)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(half_pts))


def make_blank_para(style_id=STYLE_BODY):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), style_id)
    pPr.append(pStyle)
    p.append(pPr)
    return p


def is_blue_bold(r_elem):
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        return False
    col = rPr.find(qn("w:color"))
    bold = rPr.find(qn("w:b"))
    if col is None or bold is None:
        return False
    return col.get(qn("w:val"), "").upper() == BLUE_COLOUR.upper()


# ── Core correction applier ───────────────────────────────────────────────────

def apply_corrections(docx_path, output_path, plan):
    """Apply the correction plan to the document XML directly."""

    shutil.copy2(docx_path, output_path)
    tmp = tempfile.mkdtemp()

    with zipfile.ZipFile(output_path, 'r') as z:
        z.extractall(tmp)

    doc_xml = os.path.join(tmp, "word", "document.xml")
    tree = etree.parse(doc_xml)
    root = tree.getroot()
    ns = {"w": W}
    body = root.find(".//w:body", ns)

    paragraphs = body.findall("w:p", ns)
    tables = body.findall("w:tbl", ns)

    # ── 1. Apply paragraph corrections ───────────────────────────────────────
    remove_indices = set()

    for correction in plan.get("paragraphs", []):
        idx = correction.get("index")
        action = correction.get("action", "keep")
        new_style = correction.get("new_style")
        new_text = correction.get("new_text")

        if idx is None or idx >= len(paragraphs):
            continue

        p = paragraphs[idx]

        if action == "remove":
            remove_indices.add(idx)

        elif action == "restyle" and new_style:
            set_para_style(p, new_style)
            # Fix Arial font on all runs
            for r in p.findall(".//w:r", ns):
                if not is_blue_bold(r):
                    set_run_font_arial(r)

        elif action == "fix_text" and new_text is not None:
            # Replace all text runs with the corrected text
            # Keep blue bold runs untouched
            has_blue = any(is_blue_bold(r) for r in p.findall(".//w:r", ns))
            if not has_blue:
                for r in p.findall(".//w:r", ns):
                    p.remove(r)
                r_new = OxmlElement("w:r")
                set_run_font_arial(r_new)
                t = OxmlElement("w:t")
                if new_text.startswith(" ") or new_text.endswith(" "):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = new_text
                r_new.append(t)
                p.append(r_new)

        elif action == "blank":
            # Convert to blank NormalNumberedtext paragraph
            set_para_style(p, STYLE_BODY)
            for r in p.findall(".//w:r", ns):
                p.remove(r)

    # Remove paragraphs marked for removal (in reverse order)
    for idx in sorted(remove_indices, reverse=True):
        if idx < len(paragraphs):
            body.remove(paragraphs[idx])

    # ── 2. Apply text fixes (em-dashes, doc numbers, etc.) ───────────────────
    for fix in plan.get("text_fixes", []):
        find_text = fix.get("find", "")
        replace_text = fix.get("replace", "")
        if not find_text:
            continue

        for t_elem in root.iter(qn("w:t")):
            if t_elem.text and find_text in t_elem.text:
                # Don't modify blue bold runs (placeholders)
                r_parent = t_elem.getparent()
                if r_parent is not None and is_blue_bold(r_parent):
                    continue
                t_elem.text = t_elem.text.replace(find_text, replace_text)

    # Also handle em/en dashes directly (regardless of plan)
    for t_elem in root.iter(qn("w:t")):
        if t_elem.text:
            r_parent = t_elem.getparent()
            if r_parent is not None and is_blue_bold(r_parent):
                continue
            t_elem.text = t_elem.text.replace(EM_DASH, "-").replace(EN_DASH, "-")

    # Handle API error text patterns (Python dict strings)
    dict_pattern = re.compile(r"\{['\"]text['\"]\s*:\s*['\"](.+?)['\"]\s*,\s*['\"]suggested['\"]\s*:\s*(True|False)\}", re.DOTALL)
    for t_elem in root.iter(qn("w:t")):
        if t_elem.text:
            match = dict_pattern.search(t_elem.text)
            if match:
                t_elem.text = dict_pattern.sub(match.group(1), t_elem.text)

    # Handle document number codes RTPR-HSE-XXX-NNNN
    doc_num_pattern = re.compile(r"RTPR-HSE-[A-Z]{2,5}-\d{4}")
    for t_elem in root.iter(qn("w:t")):
        if t_elem.text and doc_num_pattern.search(t_elem.text):
            r_parent = t_elem.getparent()
            if r_parent is not None and is_blue_bold(r_parent):
                continue
            t_elem.text = doc_num_pattern.sub("[INSERT: Project Document Number]", t_elem.text)

    # ── 3. Apply table corrections ────────────────────────────────────────────
    for table_correction in plan.get("tables", []):
        t_idx = table_correction.get("index")
        width = table_correction.get("width_twips", TABLE_WIDTH_H2)
        indent = table_correction.get("indent_twips", TABLE_INDENT_H2)

        if t_idx is None or t_idx >= len(tables):
            continue

        tbl = tables[t_idx]
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl.insert(0, tblPr)

        # Set table width to absolute dxa
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:type"), "dxa")
        tblW.set(qn("w:w"), str(width))

        # Set indent
        tblInd = tblPr.find(qn("w:tblInd"))
        if tblInd is None:
            tblInd = OxmlElement("w:tblInd")
            tblPr.append(tblInd)
        tblInd.set(qn("w:type"), "dxa")
        tblInd.set(qn("w:w"), str(indent))

        # Set cell margins at table level
        tblCellMar = tblPr.find(qn("w:tblCellMar"))
        if tblCellMar is None:
            tblCellMar = OxmlElement("w:tblCellMar")
            tblPr.append(tblCellMar)
        for side, val in [("top", CELL_TOP), ("bottom", CELL_BOTTOM),
                          ("left", CELL_LEFT), ("right", CELL_RIGHT)]:
            el = tblCellMar.find(qn(f"w:{side}"))
            if el is None:
                el = OxmlElement(f"w:{side}")
                tblCellMar.append(el)
            el.set(qn("w:type"), "dxa")
            el.set(qn("w:w"), str(val))

        # Fix cell widths (distribute proportionally)
        rows = tbl.findall(".//w:tr", ns)
        if rows:
            # Get column count from first row
            first_row_cells = rows[0].findall("w:tc", ns)
            col_count = len(first_row_cells) if first_row_cells else 1
            cell_width = width // col_count
            for tr in rows:
                cells = tr.findall("w:tc", ns)
                for i, tc in enumerate(cells):
                    tcPr = tc.find(qn("w:tcPr"))
                    if tcPr is None:
                        tcPr = OxmlElement("w:tcPr")
                        tc.insert(0, tcPr)
                    tcW = tcPr.find(qn("w:tcW"))
                    if tcW is None:
                        tcW = OxmlElement("w:tcW")
                        tcPr.append(tcW)
                    tcW.set(qn("w:type"), "dxa")
                    # Last cell gets remainder
                    if i == len(cells) - 1:
                        tcW.set(qn("w:w"), str(width - cell_width * (len(cells) - 1)))
                    else:
                        tcW.set(qn("w:w"), str(cell_width))

        # Fix all cell text: size 8 (sz=16), zero spacing, Arial
        for tc in tbl.findall(".//w:tc", ns):
            for p_cell in tc.findall(".//w:p", ns):
                # Zero spacing
                pPr = p_cell.find(qn("w:pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    p_cell.insert(0, pPr)
                spacing = pPr.find(qn("w:spacing"))
                if spacing is None:
                    spacing = OxmlElement("w:spacing")
                    pPr.append(spacing)
                spacing.set(qn("w:before"), "0")
                spacing.set(qn("w:after"), "0")
                spacing.set(qn("w:line"), "240")
                spacing.set(qn("w:lineRule"), "atLeast")

                # Size 8 on all runs
                for r in p_cell.findall(".//w:r", ns):
                    set_run_size(r, 16)  # 16 half-points = 8pt
                    set_run_font_arial(r)

    # ── 4. Apply traceability table corrections ───────────────────────────────
    # Find the traceability table (last table in document, or one containing "Source Requirements")
    traceability_fixes = plan.get("traceability_rows", [])
    if traceability_fixes and tables:
        # Use last table as the traceability table
        trace_table = tables[-1]
        rows = trace_table.findall(".//w:tr", ns)
        for fix in traceability_fixes:
            row_idx = fix.get("row_index")
            section_val = fix.get("section", "")
            status_val = fix.get("status", "")
            # +1 to skip header row
            actual_row = (row_idx or 0) + 1
            if actual_row < len(rows):
                row = rows[actual_row]
                cells = row.findall("w:tc", ns)
                # Column 0 = Section
                if cells and section_val:
                    for t in cells[0].iter(qn("w:t")):
                        t.text = section_val
                        break
                # Column 3 = Status (0-indexed)
                if len(cells) > 3 and status_val:
                    for t in cells[3].iter(qn("w:t")):
                        t.text = status_val
                        break

    # ── 5. Universal Arial font pass ──────────────────────────────────────────
    # Apply to all runs except blue bold (placeholders)
    # Skip front cover and TOC (first ~50 paragraphs)
    all_paras = body.findall("w:p", ns)
    toc_end_idx = min(50, len(all_paras))
    for p in all_paras[toc_end_idx:]:
        for r in p.findall(".//w:r", ns):
            if not is_blue_bold(r):
                set_run_font_arial(r)

    # Write back
    tree.write(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Rezip
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for file in files:
                fp = os.path.join(root_dir, file)
                zout.write(fp, os.path.relpath(fp, tmp))

    shutil.rmtree(tmp)
    print(f"  Saved: {output_path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",  required=True, help="Path to input .docx")
    parser.add_argument("--output", required=True, help="Path to output .docx")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"ERROR: Input file not found: {args.input}")
        sys.exit(1)

    os.makedirs(os.path.dirname(args.output), exist_ok=True)

    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    print(f"Analysing: {os.path.basename(args.input)}")
    plan = get_correction_plan(client, args.input)

    if not plan:
        print("ERROR: Could not get correction plan from Claude. Copying original.")
        shutil.copy2(args.input, args.output)
        sys.exit(1)

    para_corrections = len(plan.get("paragraphs", []))
    table_corrections = len(plan.get("tables", []))
    text_fixes = len(plan.get("text_fixes", []))
    print(f"  Plan: {para_corrections} paragraph corrections, "
          f"{table_corrections} table fixes, {text_fixes} text fixes")

    apply_corrections(args.input, args.output, plan)
    print(f"Done: {os.path.basename(args.output)}")


if __name__ == "__main__":
    main()
