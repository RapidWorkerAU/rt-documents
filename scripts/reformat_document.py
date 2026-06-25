r"""
reformat_document.py

Deterministic document reformatter. No Claude API needed.
Detects paragraph types by their run-level formatting signals and
applies the correct named styles, strips heading numbers, manages
blank lines, fixes tables, replaces em-dashes and document codes.

Detection signals (from XML inspection of generated documents):
  H1:      sz=26, colour=1F3864, bold=True  → text matches r'^\d+\s{2}'
  H2:      sz=22, colour=1F3864, bold=True  → text matches r'^\d+\.\d+\s{2}'
  Body:    sz=20, colour=2D2D2D
  Bullet:  ListParagraph style, sz=20
  Blank:   no sz, no colour, empty text
  Blue placeholder: colour=0059D1, bold=True → leave untouched
  Orange suggested: colour=C55A11 → leave text, fix style only
"""

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Target style names
STYLE_H1        = "Heading1Numbered"
STYLE_H2        = "Heading2Numbered"
STYLE_NORMAL    = "Normal"
STYLE_BODY      = "NormalNumberedtext"
STYLE_BULLET_H1 = "NumberedBulletL1"
STYLE_BULLET_H2 = "NumberedBulletL3"

# Colours
COL_HEADING = "1F3864"
COL_BODY    = "2D2D2D"
COL_BLUE    = "0059D1"
COL_ORANGE  = "C55A11"

# Table dimensions
TABLE_W_H1   = 9777
TABLE_W_H2   = 9026
TABLE_IND_H1 = 0
TABLE_IND_H2 = 704
CELL_TOP = CELL_BOTTOM = 36
CELL_LEFT = CELL_RIGHT = 108

EM_DASH = "\u2014"
EN_DASH = "\u2013"
DOC_NUM_RE = re.compile(r"RTPR-HSE-[A-Z0-9]{2,6}-\d{4}")
DICT_STR_RE = re.compile(
    r"\{['\"]text['\"]\s*:\s*['\"](.+?)['\"]\s*,\s*['\"]suggested['\"]\s*:\s*(?:True|False)\}",
    re.DOTALL
)

# Number-stripping patterns for headings
H1_NUM_RE = re.compile(r"^\d+\s{1,3}")
H2_NUM_RE = re.compile(r"^\d+\.\d+\s{1,3}")


# ── XML helpers ───────────────────────────────────────────────────────────────

def ns(tag):
    return f"{{{W}}}{tag}"


def get_style(p):
    pPr = p.find(ns("pPr"))
    if pPr is None:
        return ""
    pStyle = pPr.find(ns("pStyle"))
    if pStyle is None:
        return ""
    return pStyle.get(ns("val"), "")


def set_style(p, style_id):
    pPr = p.find(ns("pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    pStyle = pPr.find(ns("pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(ns("val"), style_id)


def get_para_text(p):
    return "".join(t.text or "" for t in p.iter(ns("t")))


def get_run_colour(r):
    rPr = r.find(ns("rPr"))
    if rPr is None:
        return ""
    col = rPr.find(ns("color"))
    if col is None:
        return ""
    return col.get(ns("val"), "").upper()


def get_run_sz(r):
    rPr = r.find(ns("rPr"))
    if rPr is None:
        return None
    sz = rPr.find(ns("sz"))
    if sz is None:
        return None
    val = sz.get(ns("val"), "")
    return int(val) if val.isdigit() else None


def is_bold(r):
    rPr = r.find(ns("rPr"))
    if rPr is None:
        return False
    return rPr.find(ns("b")) is not None


def has_blue_run(p):
    for r in p.findall(".//w:r", {"w": W}):
        if get_run_colour(r) == COL_BLUE and is_bold(r):
            return True
    return False


def para_signals(p):
    """Return (sz, colour, bold) from the dominant non-blue run."""
    runs = p.findall(".//w:r", {"w": W})
    for r in runs:
        col = get_run_colour(r)
        if col == COL_BLUE:
            continue
        sz = get_run_sz(r)
        bold = is_bold(r)
        if sz or col:
            return sz, col, bold
    return None, "", False


def is_blank_para(p):
    text = get_para_text(p).strip()
    return text == ""


def classify_para(p):
    """
    Returns one of: h1, h2, body, bullet, blank, blue_only, skip
    """
    style = get_style(p)
    text = get_para_text(p)

    # Already correctly styled headings
    if style == STYLE_H1:
        return "h1"
    if style == STYLE_H2:
        return "h2"

    # Blank
    if is_blank_para(p):
        return "blank"

    # Bullets (ListParagraph)
    if style in ("ListParagraph",):
        return "bullet"

    # Detect by run signals
    sz, col, bold = para_signals(p)

    # H1: sz=26, heading colour, bold
    if sz == 26 and col == COL_HEADING.upper() and bold:
        return "h1"

    # H2: sz=22, heading colour, bold
    if sz == 22 and col == COL_HEADING.upper() and bold:
        return "h2"

    # Blue-only placeholder paragraph
    if has_blue_run(p) and (sz == 20 or sz is None):
        all_cols = set(get_run_colour(r) for r in p.findall(".//w:r", {"w": W}))
        if all_cols <= {COL_BLUE, ""}:
            return "body"  # treat as body but preserve blue

    # Body text: sz=20
    if sz == 20:
        return "body"

    # Skip cover/TOC content
    return "skip"


def strip_heading_number(text, is_h2=False):
    """Remove leading number+spaces from heading text."""
    if is_h2:
        return H2_NUM_RE.sub("", text, count=1).strip()
    return H1_NUM_RE.sub("", text, count=1).strip()


def set_run_text(p, new_text):
    """Replace all text in a paragraph's runs with new_text, preserving first run's formatting."""
    runs = p.findall(".//w:r", {"w": W})
    if not runs:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = new_text
        r.append(t)
        p.append(r)
        return
    # Set first run text
    first_r = runs[0]
    t_elems = first_r.findall(".//w:t", {"w": W})
    if t_elems:
        t_elems[0].text = new_text
        for t in t_elems[1:]:
            t.getparent().remove(t)
    # Remove all other runs
    for r in runs[1:]:
        r.getparent().remove(r)


def set_run_font_arial(r):
    rPr = r.find(ns("rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
    fonts = rPr.find(ns("rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rPr.insert(0, fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), "Arial")


def make_blank(style=STYLE_BODY):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(ns("val"), style)
    pPr.append(pStyle)
    p.append(pPr)
    return p


def fix_text_in_run(t_elem):
    """Apply text fixes to a single w:t element."""
    if not t_elem.text:
        return
    # Em/en dashes
    t_elem.text = t_elem.text.replace(EM_DASH, "-").replace(EN_DASH, "-")
    # Document number codes
    t_elem.text = DOC_NUM_RE.sub("[INSERT: Project Document Number]", t_elem.text)
    # API error dict strings
    m = DICT_STR_RE.search(t_elem.text)
    if m:
        t_elem.text = DICT_STR_RE.sub(m.group(1), t_elem.text)
    # Round brackets around INSERT placeholders
    t_elem.text = re.sub(r'\(\s*(\[INSERT:[^\]]+\])\s*\)', r'\1', t_elem.text)


def apply_bracket_formatting(p, in_table=False):
    """
    Find any text matching [...] pattern and reformat that span as bold blue.
    Splits runs as needed to isolate bracketed spans.
    Size 10 (sz=20) in body, size 8 (sz=16) in tables.
    """
    sz_val = "16" if in_table else "20"
    bracket_re = re.compile(r'(\[[^\[\]]+\])')

    runs = list(p.findall(".//w:r", {"w": W}))
    for r in runs:
        t_elems = r.findall(ns("t"))
        if not t_elems:
            continue
        full_text = "".join(t.text or "" for t in t_elems)
        if "[" not in full_text:
            continue

        # Already blue bold — leave it
        if get_run_colour(r) == COL_BLUE and is_bold(r):
            continue

        # Split text into parts: plain and bracketed
        parts = bracket_re.split(full_text)
        if len(parts) <= 1:
            continue

        # Get parent paragraph for insertion
        parent = r.getparent()
        if parent is None:
            continue
        r_idx = list(parent).index(r)

        # Get original run rPr to copy for plain parts
        orig_rPr = r.find(ns("rPr"))

        new_runs = []
        for part in parts:
            if not part:
                continue
            new_r = OxmlElement("w:r")

            if bracket_re.match(part):
                # Bracketed span — bold blue
                rPr = OxmlElement("w:rPr")
                fonts = OxmlElement("w:rFonts")
                for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                    fonts.set(qn(attr), "Arial")
                rPr.append(fonts)
                b = OxmlElement("w:b")
                bCs = OxmlElement("w:bCs")
                rPr.append(b)
                rPr.append(bCs)
                col = OxmlElement("w:color")
                col.set(ns("val"), COL_BLUE)
                rPr.append(col)
                sz = OxmlElement("w:sz")
                sz.set(ns("val"), sz_val)
                szCs = OxmlElement("w:szCs")
                szCs.set(ns("val"), sz_val)
                rPr.append(sz)
                rPr.append(szCs)
                new_r.append(rPr)
            else:
                # Plain span — copy original formatting
                if orig_rPr is not None:
                    from copy import deepcopy
                    new_r.append(deepcopy(orig_rPr))
                else:
                    rPr = OxmlElement("w:rPr")
                    fonts = OxmlElement("w:rFonts")
                    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                        fonts.set(qn(attr), "Arial")
                    rPr.append(fonts)
                    new_r.append(rPr)

            t = OxmlElement("w:t")
            if part.startswith(" ") or part.endswith(" "):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = part
            new_r.append(t)
            new_runs.append(new_r)

        if new_runs:
            # Remove original run and insert new split runs
            parent.remove(r)
            for offset, new_r in enumerate(new_runs):
                parent.insert(r_idx + offset, new_r)


# ── Table fixer ───────────────────────────────────────────────────────────────

def fix_table(tbl, context):
    """Fix table widths, indents, cell margins, cell text size and spacing."""
    width  = TABLE_W_H1 if context == "h1" else TABLE_W_H2
    indent = TABLE_IND_H1 if context == "h1" else TABLE_IND_H2

    tblPr = tbl.find(ns("tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # Table width
    tblW = tblPr.find(ns("tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(ns("type"), "dxa")
    tblW.set(ns("w"), str(width))

    # Table indent
    tblInd = tblPr.find(ns("tblInd"))
    if tblInd is None:
        tblInd = OxmlElement("w:tblInd")
        tblPr.append(tblInd)
    tblInd.set(ns("type"), "dxa")
    tblInd.set(ns("w"), str(indent))

    # Cell margins
    tblCellMar = tblPr.find(ns("tblCellMar"))
    if tblCellMar is None:
        tblCellMar = OxmlElement("w:tblCellMar")
        tblPr.append(tblCellMar)
    for side, val in [("top", CELL_TOP), ("bottom", CELL_BOTTOM),
                      ("left", CELL_LEFT), ("right", CELL_RIGHT)]:
        el = tblCellMar.find(ns(side))
        if el is None:
            el = OxmlElement(f"w:{side}")
            tblCellMar.append(el)
        el.set(ns("type"), "dxa")
        el.set(ns("w"), str(val))

    # Column widths
    rows = tbl.findall(".//w:tr", {"w": W})
    if rows:
        first_cells = rows[0].findall("w:tc", {"w": W})
        col_count = max(len(first_cells), 1)
        cell_w = width // col_count
        for tr in rows:
            cells = tr.findall("w:tc", {"w": W})
            for i, tc in enumerate(cells):
                # Skip merged cells - check for gridSpan
                tcPr = tc.find(ns("tcPr"))
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    tc.insert(0, tcPr)
                # Remove cell-level margin overrides
                for side in ["top", "bottom", "left", "right"]:
                    existing = tcPr.find(f"w:tcMar/w:{side}", {"w": W})
                    # Leave merged cell handling alone for now
                tcW = tcPr.find(ns("tcW"))
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(ns("type"), "dxa")
                if i == len(cells) - 1:
                    tcW.set(ns("w"), str(width - cell_w * (len(cells) - 1)))
                else:
                    tcW.set(ns("w"), str(cell_w))

    # Cell text: size 8 (sz=16), zero spacing, Arial
    for tc in tbl.findall(".//w:tc", {"w": W}):
        # Remove leading/trailing blank paragraphs in cells
        cell_paras = tc.findall("w:p", {"w": W})
        if len(cell_paras) > 1:
            if is_blank_para(cell_paras[0]):
                tc.remove(cell_paras[0])
                cell_paras = cell_paras[1:]
            if len(cell_paras) > 1 and is_blank_para(cell_paras[-1]):
                tc.remove(cell_paras[-1])

        for p_cell in tc.findall(".//w:p", {"w": W}):
            # Zero spacing
            pPr = p_cell.find(ns("pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p_cell.insert(0, pPr)
            spacing = pPr.find(ns("spacing"))
            if spacing is None:
                spacing = OxmlElement("w:spacing")
                pPr.append(spacing)
            spacing.set(ns("before"), "0")
            spacing.set(ns("after"), "0")
            spacing.set(ns("line"), "240")
            spacing.set(ns("lineRule"), "atLeast")

            # Size 8 on all runs, Arial
            for r in p_cell.findall(".//w:r", {"w": W}):
                set_run_font_arial(r)
                rPr = r.find(ns("rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    r.insert(0, rPr)
                for sz_tag in (ns("sz"), ns("szCs")):
                    sz_el = rPr.find(sz_tag)
                    if sz_el is None:
                        sz_el = OxmlElement(sz_tag.replace(f"{{{W}}}", "w:"))
                        rPr.append(sz_el)
                    sz_el.set(ns("val"), "16")

            # Fix text in cells
            for t_el in p_cell.findall(".//w:t", {"w": W}):
                r_parent = t_el.getparent()
                if r_parent is not None:
                    col = get_run_colour(r_parent)
                    if col == COL_BLUE:
                        continue
                fix_text_in_run(t_el)
            # Apply bracket formatting in table cells
            apply_bracket_formatting(p_cell, in_table=True)


# ── Main document fixer ───────────────────────────────────────────────────────

def find_body_start(paragraphs):
    """
    Find the index of the first body paragraph after the TOC.
    The TOC section ends after we see section heading "1" or the first H1 after cover.
    We look for the last TOC/cover-related paragraph and return the next one.
    """
    # Cover is paragraphs 0-2 (Subtitle, FrontcoverHeading, AutomatedProjectName)
    # Then come blank paragraphs, table rows for doc control, then TOC
    # Body starts after the TOC which contains a paragraph with "Contents" text
    # and then the TOC field entries ending with the last heading entry
    # Safest: find first H1 (sz=26, navy, bold) and start from 2 before it
    for i, p in enumerate(paragraphs):
        sz, col, bold = para_signals(p)
        if sz == 26 and col == COL_HEADING.upper() and bold:
            # Check it's not a TOC entry (TOC entries have PAGEREF in them)
            text = get_para_text(p)
            if "PAGEREF" not in text and "TOC" not in text:
                # Start from this paragraph (body starts here)
                return i
    return 0


def reformat(docx_path, output_path):
    """Main reformatting function."""
    shutil.copy2(docx_path, output_path)
    tmp = tempfile.mkdtemp()

    with zipfile.ZipFile(output_path, 'r') as z:
        z.extractall(tmp)

    doc_xml = os.path.join(tmp, "word", "document.xml")
    tree = etree.parse(doc_xml)
    root = tree.getroot()
    body = root.find(f"{{{W}}}body")
    ns_map = {"w": W}

    # Get all top-level body elements
    all_paras = body.findall("w:p", ns_map)
    all_tables = body.findall("w:tbl", ns_map)

    body_start = find_body_start(all_paras)
    print(f"  Body starts at paragraph {body_start}")

    # ── Pass 1: Classify and restyle all body paragraphs ─────────────────────
    last_heading_level = 1  # track whether we're under H1 or H2
    new_body = []           # (para_element, classification)

    for i, p in enumerate(all_paras):
        if i < body_start:
            new_body.append((p, "skip"))
            continue

        kind = classify_para(p)

        if kind == "h1":
            last_heading_level = 1
            text = get_para_text(p)
            stripped = strip_heading_number(text, is_h2=False)
            if stripped != text:
                set_run_text(p, stripped)
            set_style(p, STYLE_H1)
            # Ensure numPr ilvl=0, numId=1 for Heading1Numbered
            pPr = p.find(ns("pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p.insert(0, pPr)
            numPr = pPr.find(ns("numPr"))
            if numPr is None:
                numPr = OxmlElement("w:numPr")
                pPr.append(numPr)
            ilvl_el = numPr.find(ns("ilvl"))
            if ilvl_el is None:
                ilvl_el = OxmlElement("w:ilvl")
                numPr.insert(0, ilvl_el)
            ilvl_el.set(ns("val"), "0")
            numId_el = numPr.find(ns("numId"))
            if numId_el is None:
                numId_el = OxmlElement("w:numId")
                numPr.append(numId_el)
            numId_el.set(ns("val"), "1")
            for r in p.findall(".//w:r", {"w": W}):
                set_run_font_arial(r)
            new_body.append((p, "h1"))

        elif kind == "h2":
            last_heading_level = 2
            text = get_para_text(p)
            stripped = strip_heading_number(text, is_h2=True)
            if stripped != text:
                set_run_text(p, stripped)
            set_style(p, STYLE_H2)
            # Ensure numPr ilvl=1 for Heading2Numbered
            pPr = p.find(ns("pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p.insert(0, pPr)
            numPr = pPr.find(ns("numPr"))
            if numPr is None:
                numPr = OxmlElement("w:numPr")
                pPr.append(numPr)
            ilvl_el = numPr.find(ns("ilvl"))
            if ilvl_el is None:
                ilvl_el = OxmlElement("w:ilvl")
                numPr.insert(0, ilvl_el)
            ilvl_el.set(ns("val"), "1")
            numId_el = numPr.find(ns("numId"))
            if numId_el is None:
                numId_el = OxmlElement("w:numId")
                numPr.append(numId_el)
            numId_el.set(ns("val"), "1")
            for r in p.findall(".//w:r", {"w": W}):
                set_run_font_arial(r)
            new_body.append((p, "h2"))

        elif kind == "body":
            # Assign correct body style based on heading context
            if last_heading_level == 1:
                set_style(p, STYLE_NORMAL)
            else:
                set_style(p, STYLE_BODY)
            # Arial on non-blue runs
            for r in p.findall(".//w:r", {"w": W}):
                if get_run_colour(r) != COL_BLUE:
                    set_run_font_arial(r)
            # Fix text
            for t_el in p.findall(".//w:t", {"w": W}):
                r_parent = t_el.getparent()
                if r_parent is not None and get_run_colour(r_parent) == COL_BLUE:
                    continue
                fix_text_in_run(t_el)
            # Apply bracket formatting
            apply_bracket_formatting(p, in_table=False)
            new_body.append((p, "body"))

        elif kind == "bullet":
            # Under H1: NumberedBulletL1 — style definition handles numPr (ilvl=3, numId=2)
            # Under H2: NumberedBulletL3 — explicit ilvl=0, numId=2, left=1474, hanging=227
            if last_heading_level == 1:
                set_style(p, STYLE_BULLET_H1)
                # Remove any explicit numPr — let the style definition handle it
                pPr = p.find(ns("pPr"))
                if pPr is not None:
                    numPr = pPr.find(ns("numPr"))
                    if numPr is not None:
                        pPr.remove(numPr)
            else:
                set_style(p, STYLE_BULLET_H2)
                # Set explicit numPr for NumberedBulletL3
                pPr = p.find(ns("pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    p.insert(0, pPr)
                # Remove existing numPr and replace cleanly
                existing_numPr = pPr.find(ns("numPr"))
                if existing_numPr is not None:
                    pPr.remove(existing_numPr)
                numPr = OxmlElement("w:numPr")
                ilvl_el = OxmlElement("w:ilvl")
                ilvl_el.set(ns("val"), "0")
                numId_el = OxmlElement("w:numId")
                numId_el.set(ns("val"), "2")
                numPr.append(ilvl_el)
                numPr.append(numId_el)
                pPr.append(numPr)
                # Set explicit indent matching NumberedBulletL3
                ind_el = pPr.find(ns("ind"))
                if ind_el is None:
                    ind_el = OxmlElement("w:ind")
                    pPr.append(ind_el)
                ind_el.set(ns("left"), "1474")
                ind_el.set(ns("hanging"), "227")
            for r in p.findall(".//w:r", {"w": W}):
                if get_run_colour(r) != COL_BLUE:
                    set_run_font_arial(r)
            for t_el in p.findall(".//w:t", {"w": W}):
                r_parent = t_el.getparent()
                if r_parent is not None and get_run_colour(r_parent) == COL_BLUE:
                    continue
                fix_text_in_run(t_el)
            apply_bracket_formatting(p, in_table=False)
            new_body.append((p, "bullet"))

        elif kind == "blank":
            # Convert blank to NormalNumberedtext style
            set_style(p, STYLE_BODY)
            new_body.append((p, "blank"))

        else:
            new_body.append((p, "skip"))

    # ── Pass 1b: Remove consecutive duplicate paragraphs ───────────────────────
    seen_texts = []
    for i in range(len(new_body) - 1, -1, -1):
        p, kind = new_body[i]
        if kind in ("skip", "blank"):
            continue
        text = get_para_text(p).strip()
        # Only dedup paragraphs with meaningful content (>3 chars)
        if not text or len(text) <= 3:
            continue
        # Check if same text appears in the next non-blank, non-skip paragraph
        for j in range(i + 1, len(new_body)):
            p2, kind2 = new_body[j]
            if kind2 in ("skip", "blank"):
                continue
            text2 = get_para_text(p2).strip()
            if text2 == text:
                # Remove the duplicate (earlier one — keep last occurrence)
                body.remove(p)
                new_body[i] = (p, "removed")
            break

    # Clean removed entries
    new_body = [(p, k) for p, k in new_body if k != "removed"]

    # ── Pass 2: Fix blank line management ────────────────────────────────────
    # Build the corrected sequence of elements
    # Rules:
    # - Exactly one blank before H1 (except first body element)
    # - Exactly one blank before H2
    # - Exactly one blank before every table
    # - Exactly one blank after every table
    # - One blank after last bullet in a block
    # - One blank between consecutive body paragraphs
    # - Never two consecutive blanks
    # - No blank between lead-in sentence (ends with :) and first bullet

    # Get all body children in order (paras and tables interleaved)
    body_children = list(body)
    body_start_elem = all_paras[body_start] if body_start < len(all_paras) else None

    # Rebuild body content from body_start onwards
    pre_body = []
    post_body = []
    in_body = False

    for child in body_children:
        if not in_body:
            if child is body_start_elem:
                in_body = True
                post_body.append(child)
            else:
                pre_body.append(child)
        else:
            post_body.append(child)

    # sectPr is always last
    sectPr = body.find("w:sectPr", ns_map)

    # Create classification map for the post_body elements
    para_class = {id(p): kind for p, kind in new_body}

    def get_class(elem):
        if elem.tag == ns("tbl"):
            return "table"
        return para_class.get(id(elem), "skip")

    # Rebuild with correct blanks
    result = []
    prev_class = "none"
    prev_text = ""

    for elem in post_body:
        if elem.tag == ns("sectPr"):
            continue
        curr_class = get_class(elem)

        if curr_class == "blank":
            # Don't add yet — we'll insert blanks ourselves
            continue

        # Determine if we need a blank before this element
        need_blank = False

        if curr_class == "h1":
            need_blank = (prev_class != "none")
        elif curr_class == "h2":
            need_blank = True
        elif curr_class == "table":
            need_blank = True
        elif curr_class == "body":
            # One blank between consecutive body paras
            if prev_class in ("body",):
                need_blank = True
            elif prev_class == "bullet":
                need_blank = True
        elif curr_class == "bullet":
            if prev_class == "body":
                # Only blank if previous body para does NOT end with ':'
                if not prev_text.rstrip().endswith(":"):
                    need_blank = True
            elif prev_class == "table":
                need_blank = True
            # No blank between consecutive bullets

        # Remove trailing blank from result if it would create a double blank
        # then add a blank if needed
        if need_blank:
            # Remove consecutive blanks at end of result
            while result and result[-1][0] == "blank_spacer":
                result.pop()
            result.append(("blank_spacer", make_blank()))

        result.append(("elem", elem))
        prev_class = curr_class
        if elem.tag != ns("tbl"):
            prev_text = get_para_text(elem)

        # After a table, always add a blank
        if curr_class == "table":
            result.append(("blank_spacer", make_blank()))
            prev_class = "blank"

    # Rebuild body
    # Remove all existing children from body
    for child in list(body):
        body.remove(child)

    # Add pre-body elements
    for child in pre_body:
        body.append(child)

    # Add corrected body
    for kind, elem in result:
        body.append(elem)

    # Add sectPr back
    if sectPr is not None:
        body.append(sectPr)

    # ── Pass 3: Fix tables ────────────────────────────────────────────────────
    # Re-fetch tables after DOM rebuild
    all_tables_new = body.findall("w:tbl", ns_map)

    # Determine context (under H1 or H2) for each table
    for tbl in all_tables_new:
        # Look backwards for the nearest heading
        prev_sib = tbl.getprevious()
        context = "h2"  # default
        while prev_sib is not None:
            if prev_sib.tag == ns("p"):
                sib_style = get_style(prev_sib)
                if sib_style == STYLE_H1:
                    context = "h1"
                    break
                elif sib_style == STYLE_H2:
                    context = "h2"
                    break
                elif sib_style not in (STYLE_BODY, STYLE_NORMAL, ""):
                    pass
            prev_sib = prev_sib.getprevious()
        fix_table(tbl, context)

    # ── Pass 4: Fix text in body paragraphs we may have missed ───────────────
    for p in body.findall(".//w:p", ns_map):
        for t_el in p.findall(".//w:t", {"w": W}):
            r_parent = t_el.getparent()
            if r_parent is not None and get_run_colour(r_parent) == COL_BLUE:
                continue
            fix_text_in_run(t_el)

    # ── Write back ────────────────────────────────────────────────────────────
    tree.write(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for file in files:
                fp = os.path.join(root_dir, file)
                zout.write(fp, os.path.relpath(fp, tmp))

    shutil.rmtree(tmp)
    print(f"  Saved: {output_path}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",  required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"ERROR: {args.input} not found")
        sys.exit(1)

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    print(f"Reformatting: {os.path.basename(args.input)}")
    reformat(args.input, args.output)
    print("Done.")


if __name__ == "__main__":
    main()
