"""
generate_procedure.py

Generates a procedure document by injecting Claude-generated content into
the procedure_template.docx. All paragraphs use the template's named Word
styles (Heading1Numbered, Heading2Numbered, NormalNumberedtext, etc.) so
the output is styled identically to the example procedure.
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile

import anthropic
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# ── Colour hex values (for inline colour overrides only) ─────────────────────
BLUE   = "0059D1"   # INSERT placeholders
ORANGE = "C55A11"   # Suggested content
WHITE  = "FFFFFF"
NAVY   = "1F3864"   # Table headers


# ── Simple paragraph builder using template styles ───────────────────────────

def styled_para(style_id, text="", colour=None, bold=False):
    """Build a <w:p> using a named template style. Colour/bold only when overriding."""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), style_id)
    pPr.append(ps)
    p.append(pPr)
    if text:
        r = OxmlElement("w:r")
        if colour or bold:
            rPr = OxmlElement("w:rPr")
            if bold:
                rPr.append(OxmlElement("w:b"))
                rPr.append(OxmlElement("w:bCs"))
            if colour:
                c = OxmlElement("w:color"); c.set(qn("w:val"), colour)
                rPr.append(c)
            r.append(rPr)
        t = OxmlElement("w:t")
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)
        p.append(r)
    return p


def multi_run_para(style_id, runs):
    """
    Build a paragraph with multiple runs of different formatting.
    runs = list of (text, colour, bold) tuples
    """
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), style_id)
    pPr.append(ps)
    p.append(pPr)
    for text, colour, bold in runs:
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        if bold:
            rPr.append(OxmlElement("w:b"))
            rPr.append(OxmlElement("w:bCs"))
        if colour:
            c = OxmlElement("w:color"); c.set(qn("w:val"), colour)
            rPr.append(c)
        r.append(rPr)
        t = OxmlElement("w:t")
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)
        p.append(r)
    return p


def blank():
    """Empty paragraph using Default style for spacing."""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), "Default")
    pPr.append(ps)
    p.append(pPr)
    return p


# ── Style shortcuts ───────────────────────────────────────────────────────────

def h1(text):
    return [styled_para("Heading1Numbered", text)]

def h2(text):
    return [styled_para("Heading2Numbered", text)]

def h3(text):
    return [styled_para("Heading3Numbered", text)]

def body(text, colour=None):
    return [styled_para("NormalNumberedtext", text, colour=colour)]

def bullet(text, colour=None):
    return [styled_para("NumberedBulletL1", text, colour=colour)]

def insert_ph(field, hint=""):
    text = f"[INSERT: {field}]"
    if hint:
        text += f" — {hint}"
    return [styled_para("NormalNumberedtext", text, colour=BLUE, bold=True)]

def table_text(text, colour=None, bold=False):
    return styled_para("TableContents", text, colour=colour, bold=bold)

def table_bullet(text, colour=None):
    return styled_para("TableBulletPoints", text, colour=colour)


# ── Shaded instruction/notice box (single-cell table) ────────────────────────

def shaded_box(text, fill, text_colour=None, bold=False):
    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblStyle = OxmlElement("w:tblStyle"); tblStyle.set(qn("w:val"), "TableOption1")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:type"), "pct"); tblW.set(qn("w:w"), "5000")
    tblPr.append(tblStyle); tblPr.append(tblW)
    tbl.append(tblPr)

    tr = OxmlElement("w:tr")
    tc = OxmlElement("w:tc")
    tcPr = OxmlElement("w:tcPr")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd); tc.append(tcPr)
    p = styled_para("Instructions", text, colour=text_colour, bold=bold)
    tc.append(p); tr.append(tc); tbl.append(tr)
    return [tbl, blank()]


# ── Full table builder ────────────────────────────────────────────────────────

def make_table(headers, rows, suggested=False):
    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblStyle = OxmlElement("w:tblStyle"); tblStyle.set(qn("w:val"), "TableOption2")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:type"), "pct"); tblW.set(qn("w:w"), "5000")
    tblPr.append(tblStyle); tblPr.append(tblW)
    tbl.append(tblPr)

    def make_cell(text, fill=None, is_header=False):
        tc = OxmlElement("w:tc")
        tcPr = OxmlElement("w:tcPr")
        if fill:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
            tcPr.append(shd)
        tc.append(tcPr)
        val = str(text)
        is_insert = val.startswith("[INSERT:")
        is_orange = suggested and not is_insert and not is_header
        colour = BLUE if is_insert else (ORANGE if is_orange else (WHITE if is_header else None))
        bold = is_insert or is_header
        if is_header:
            p = table_text(val, colour=WHITE, bold=True)
        else:
            p = table_text(val, colour=colour, bold=bold)
        tc.append(p)
        return tc

    # Header row
    hdr = OxmlElement("w:tr")
    trPr = OxmlElement("w:trPr")
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader); hdr.append(trPr)
    for h in headers:
        hdr.append(make_cell(h, fill=NAVY, is_header=True))
    tbl.append(hdr)

    for row in rows:
        tr = OxmlElement("w:tr")
        for cell_val in row:
            tr.append(make_cell(cell_val))
        tbl.append(tr)

    return [tbl, blank()]


# ── Section 0 ─────────────────────────────────────────────────────────────────

def build_section_0(brief):
    elems = []
    elems += shaded_box(
        "INSTRUCTION: Read, action and delete this box before this procedure is issued. "
        "Delete the entire Section 0 before this document is submitted for review or approval. "
        "It must not appear in any approved version of this procedure.",
        fill="FFF2CC", text_colour=ORANGE, bold=True
    )
    elems += h1("0  How to Use This Document")
    elems += body(
        "This section explains the purpose of the template, how it is structured, the sequence "
        "for completing it and the checks required before it is issued. It must be read in full "
        "before any section of this document is completed. Delete Section 0 in its entirety before issue."
    )
    elems += h2("0.1  Purpose of This Template")
    elems += body(
        f"This template supports the preparation of a project-specific {brief['title']} for projects "
        "delivered under an Owner-Led model. Under this model, Rio Tinto directly manages all project "
        "activities, and this procedure applies to all personnel on the project, including Rio Tinto "
        "staff and all contractor organisations."
    )
    elems += body(
        "This template is provided as a 90% draft. Fixed content reflects the minimum Rio Tinto "
        "requirement and must not be reduced. The remaining 10% covers role names, register locations, "
        "schedules and project-specific triggers, and must be completed by the project team before issue."
    )
    elems += h2("0.2  How This Document is Structured")
    elems += body("This document contains four content types:")

    # Content type descriptions with mixed formatting
    elems.append(multi_run_para("NormalNumberedtext", [
        ("Fixed content ", None, True),
        ("is plain black text reflecting minimum Rio Tinto requirements. Fixed content must not be reduced, deleted or overridden.", None, False),
    ]))
    elems.append(multi_run_para("NormalNumberedtext", [
        ("[INSERT: placeholder text] ", BLUE, True),
        ("appears in blue bold. Each placeholder must be replaced with accurate, project-specific information before issue.", None, False),
    ]))
    elems.append(multi_run_para("NormalNumberedtext", [
        ("Orange text ", ORANGE, False),
        ("marks specific who, what, when, or how detail included as a best practice suggestion. The project team may retain, modify, or replace orange detail.", None, False),
    ]))

    elems += h2("0.3  Pre-Issue Checklist")
    elems += body(
        "Before submitting this document for review or approval, confirm all items below are complete, "
        "then delete this checklist together with the rest of Section 0."
    )
    checklist = [
        ("Cover page is complete including document number, revision, date and signatories", "☐"),
        ("All [INSERT:] placeholder text has been replaced with project-specific information", "☐"),
        ("All instruction boxes have been read, actioned and deleted", "☐"),
        ("All role titles match the actual project organisational structure", "☐"),
        ("All system and register locations are confirmed and inserted", "☐"),
        ("All workshop and meeting dates are confirmed against the project programme", "☐"),
        ("Related procedures in the References section are issued or being issued concurrently", "☐"),
        ("Section 17 (Source Requirements Traceability) will be deleted before final issue", "☐"),
        ("Section 16 (Project Localisation Edits) has been populated with all changes made during preparation", "☐"),
        ("Section 0 has been deleted in its entirety", "☐"),
    ]
    elems += make_table(["Check", "Complete?"], checklist)
    return elems


# ── Roles table ───────────────────────────────────────────────────────────────

def build_roles_table():
    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblStyle = OxmlElement("w:tblStyle"); tblStyle.set(qn("w:val"), "TableOption2")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:type"), "pct"); tblW.set(qn("w:w"), "5000")
    tblPr.append(tblStyle); tblPr.append(tblW)
    tbl.append(tblPr)

    def hdr_cell(text):
        tc = OxmlElement("w:tc")
        tcPr = OxmlElement("w:tcPr")
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), NAVY)
        tcPr.append(shd); tc.append(tcPr)
        tc.append(table_text(text, colour=WHITE, bold=True))
        return tc

    def group_cell(text):
        """Full-width role group header row."""
        tc = OxmlElement("w:tc")
        tcPr = OxmlElement("w:tcPr")
        gs = OxmlElement("w:gridSpan"); gs.set(qn("w:val"), "4")
        tcPr.append(gs); tc.append(tcPr)
        tc.append(styled_para("TableHeading", text))
        return tc

    def resp_cell(text):
        tc = OxmlElement("w:tc"); tc.append(OxmlElement("w:tcPr"))
        tc.append(table_bullet(text))
        return tc

    def tick_cell(tick):
        tc = OxmlElement("w:tc"); tc.append(OxmlElement("w:tcPr"))
        tc.append(table_text(tick, bold=bool(tick)))
        return tc

    def insert_row():
        tr = OxmlElement("w:tr")
        tc = OxmlElement("w:tc")
        tcPr = OxmlElement("w:tcPr")
        gs = OxmlElement("w:gridSpan"); gs.set(qn("w:val"), "4")
        tcPr.append(gs); tc.append(tcPr)
        tc.append(table_text("[INSERT: add further responsibilities as required]", colour=BLUE, bold=True))
        tr.append(tc)
        return tr

    # Header
    hdr = OxmlElement("w:tr")
    trPr = OxmlElement("w:trPr")
    tblH = OxmlElement("w:tblHeader"); trPr.append(tblH); hdr.append(trPr)
    for h_text in ["Responsibilities", "Company", "Contractor", "Section"]:
        hdr.append(hdr_cell(h_text))
    tbl.append(hdr)

    role_groups = [
        ("Project Management", [
            ("Holds overall accountability for project compliance with this procedure and approves this procedure and any material amendment.", "✓", ""),
            ("Accepts residual risk at the highest risk class level and authorises continuation where escalation is required.", "✓", ""),
        ]),
        ("HSE/CSP Management", [
            ("Maintains and implements this procedure and confirms all activities meet its requirements before work proceeds.", "✓", "✓"),
            ("Reviews and approves all required assessments, plans, and registers produced under this procedure.", "✓", "✓"),
            ("Reports performance against this procedure at the agreed project HSESC performance meeting.", "✓", "✓"),
            ("Reviews performance data, identifies trends, and incorporates systemic findings into corrective action programmes.", "✓", "✓"),
        ]),
        ("Frontline Supervision", [
            ("Ensures work crew compliance with this procedure before and during all relevant task executions.", "✓", "✓"),
            ("Stops work when the requirements of this procedure cannot be met and notifies the HSESC Manager.", "✓", "✓"),
            ("Confirms all personnel under their supervision are trained and competent before commencing relevant activities.", "✓", "✓"),
        ]),
        ("All Project Personnel / Workers", [
            ("Complies with the requirements of this procedure in all relevant work activities.", "✓", "✓"),
            ("Reports any non-compliance with this procedure immediately to their supervisor.", "✓", "✓"),
        ]),
    ]

    for role_name, responsibilities in role_groups:
        grp = OxmlElement("w:tr")
        grp.append(group_cell(role_name))
        tbl.append(grp)
        for resp_text, company, contractor in responsibilities:
            r = OxmlElement("w:tr")
            r.append(resp_cell(resp_text))
            r.append(tick_cell(company))
            r.append(tick_cell(contractor))
            r.append(tick_cell(""))
            tbl.append(r)
        tbl.append(insert_row())

    return [tbl, blank()]


# ── Content item renderer ─────────────────────────────────────────────────────

def render_item(item):
    if isinstance(item, str):
        return body(item)

    t = item.get("type", "paragraph")

    if t == "paragraph":
        colour = ORANGE if item.get("suggested") else None
        return body(item.get("text", ""), colour=colour)

    elif t == "insert":
        return insert_ph(item.get("field", "value"), item.get("hint", ""))

    elif t == "bullet_list":
        elems = []
        lead = item.get("lead", "")
        if lead:
            elems += body(lead)
        for b in item.get("items", []):
            if isinstance(b, dict):
                if b.get("insert"):
                    elems += [styled_para("NumberedBulletL1",
                        f"[INSERT: {b.get('field','value')}]", colour=BLUE, bold=True)]
                else:
                    colour = ORANGE if b.get("suggested") else None
                    elems += bullet(b.get("text", ""), colour=colour)
            else:
                elems += bullet(str(b))
        return elems

    elif t == "table":
        elems = []
        caption = item.get("caption", "")
        if caption:
            elems += body(caption)
        headers = item.get("headers", [])
        rows = item.get("rows", [])
        if headers:
            elems += make_table(headers, rows, suggested=item.get("suggested", False))
        return elems

    elif t == "instruction_box":
        return shaded_box(item.get("text", ""), fill="FFF2CC", text_colour=ORANGE, bold=True)

    return []


def build_content_sections(sections_data):
    elems = []
    for i, sec in enumerate(sections_data):
        sec_num = i + 3
        title = sec.get("title", "")
        intro = sec.get("intro", "")

        elems += h1(f"{sec_num}  {title}")
        if intro:
            elems += body(intro)

        sub_num = 1
        for sub in sec.get("subsections", []):
            sub_title = sub.get("title", "")
            sub_intro = sub.get("intro", "")

            if sub_title:
                elems += h2(f"{sec_num}.{sub_num}  {sub_title}")
            if sub_intro:
                elems += body(sub_intro)

            for l3 in sub.get("level3", []):
                l3_title = l3.get("title", "")
                l3_num = l3.get("number", sub_num)
                if l3_title:
                    elems += h3(f"{sec_num}.{sub_num}.{l3_num}  {l3_title}")
                for item in l3.get("content", []):
                    elems += render_item(item)

            for item in sub.get("content", []):
                elems += render_item(item)

            if sub_title:
                sub_num += 1

    return elems, 3 + len(sections_data)


def build_definitions(n, definitions):
    elems = []
    elems += h1(f"{n}  Definitions and Acronyms")
    elems += body(
        "The following terms are used throughout this procedure. Where a term is also used in other "
        "project HSESC procedures, the definition below applies consistently across the system."
    )
    rows = [(d.get("term", ""), d.get("definition", "")) for d in definitions]
    rows.append(("[INSERT: add project-specific terms]", "[INSERT: definition]"))
    elems += make_table(["Term", "Definition"], rows)
    return elems, n + 1


def build_references(n, references):
    elems = []
    elems += h1(f"{n}  References and Source Information")
    elems += body(
        "This section lists the internal and external documents this procedure relies on. "
        "This section remains in the procedure once published. Where a referenced document is updated, "
        "this procedure is reviewed to confirm the reference remains current."
    )
    rows = [(r.get("document", ""), r.get("reference", ""), r.get("type", "")) for r in references]
    rows.append(("[INSERT: further referenced document, where applicable]", "[INSERT: reference]", "[INSERT: type]"))
    elems += make_table(["Document", "Reference", "Type"], rows)
    return elems, n + 1


def build_localisation_edits(n):
    elems = []
    elems += h1(f"{n}  Project Localisation Edits")
    elems += body(
        "This section captures the project-specific edits made to this document during preparation. "
        "This section stays in the document once published and is updated at each document revision. "
        "Record all additions, deletions, and wording changes from the original template."
    )
    for heading_text in [
        "Records of wording changes",
        "Records of new content additions",
        "Records of content removal",
    ]:
        elems += body(heading_text)
        rows = [("[INSERT: description of change]", "[INSERT: section]")] * 3
        elems += make_table(["Description of the Change", "Section"], rows)
    return elems, n + 1


def build_traceability(n, brief):
    elems = []
    elems += shaded_box(
        "REMOVABLE SECTION — Source Requirements Traceability. Delete this entire section before "
        "the document is issued for approval. It must not appear in any approved version of this document. "
        "Table A records requirements addressed in this procedure. "
        "Table B records requirements not fully addressed, with written justification for each.",
        fill="FCE4D6", text_colour=ORANGE, bold=True
    )
    elems += h1(f"{n}  Source Requirements Traceability (DELETE BEFORE ISSUE)")

    elems += h2(f"{n}.1  Requirements Addressed in This Procedure")
    rows_a = [
        ("[INSERT: section]",
         f"{src.get('framework','')} {src.get('clause','')}".strip(),
         src.get("summary", ""),
         "Full")
        for src in brief.get("external_sources", [])
    ]
    if not rows_a:
        rows_a = [("[INSERT: section]", "[INSERT: source]", "[INSERT: requirement]", "[INSERT: Full/Partial]")]
    elems += make_table(["Section", "Source", "Requirement", "Status"], rows_a)

    elems += h2(f"{n}.2  Requirements Not Fully Addressed in This Procedure")
    elems += body(
        "List any requirements from the external sources above that are not fully addressed in this "
        "document, with a specific written justification for each. Acceptable justifications: "
        "(a) the requirement is owned by a named standalone procedure cross-referenced here; "
        "(b) excluded by a specific business decision endorsed by project leadership, named here; "
        "or (c) addressed through an operational control outside this document, identified here."
    )
    elems += make_table(
        ["Requirement Not Covered", "Source", "Justification"],
        [("[INSERT: description of requirement not covered]",
          "[INSERT: source and clause]",
          "[INSERT: specific written justification]")]
    )
    return elems


# ── Template injection ────────────────────────────────────────────────────────

def inject_into_template(template_path, output_path, brief, all_elements):
    shutil.copy2(template_path, output_path)
    tmp = tempfile.mkdtemp()

    with zipfile.ZipFile(output_path, 'r') as z:
        z.extractall(tmp)

    doc_xml = os.path.join(tmp, "word", "document.xml")
    tree = etree.parse(doc_xml)
    root = tree.getroot()
    ns = {"w": W}
    body_elem = root.find(".//w:body", ns)

    # Replace PROCEDURE TITLE
    proc_title = brief["title"].replace(" Procedure", "").replace(" Programme", "")
    for t in root.iter(qn("w:t")):
        if t.text and "PROCEDURE TITLE" in t.text:
            t.text = t.text.replace("PROCEDURE TITLE", proc_title)

    # Find and replace the body content placeholder
    placeholder = None
    for p in body_elem.findall("w:p", ns):
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if "Enter Body Content Text" in text:
            placeholder = p
            break

    if placeholder is not None:
        idx = list(body_elem).index(placeholder)
        body_elem.remove(placeholder)
    else:
        # Fallback: insert before sectPr
        sectPr = body_elem.find("w:sectPr", ns)
        idx = list(body_elem).index(sectPr) if sectPr is not None else len(list(body_elem))

    for i, elem in enumerate(all_elements):
        body_elem.insert(idx + i, elem)

    tree.write(doc_xml, xml_declaration=True, encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for file in files:
                fp = os.path.join(root_dir, file)
                zout.write(fp, os.path.relpath(fp, tmp))

    shutil.rmtree(tmp)
    print(f"  Written on template: {output_path}")


# ── Claude generation ─────────────────────────────────────────────────────────

def _call_claude(client, system, prompt, label=""):
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=16000,
                system=system,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text.strip()
            raw = re.sub(r"^```(?:json)?\s*", "", raw)
            raw = re.sub(r"\s*```$", "", raw)
            return json.loads(raw)
        except json.JSONDecodeError as e:
            print(f"  JSON error {label} attempt {attempt+1}: {e}")
            if attempt < 2:
                time.sleep(5)
            else:
                raise
        except Exception as e:
            print(f"  API error {label} attempt {attempt+1}: {e}")
            if attempt < 2:
                time.sleep(10)
            else:
                raise


def generate_content(client, brief, example_text):
    requirements_text = "\n".join(
        f"- [{r.get('tag','RT Standard')}] under '{r.get('section_heading','')}': {r.get('text','')}"
        for r in brief.get("requirements", [])
    )
    pcg_text = "\n".join(
        f"{i+1}. {item}"
        for i, item in enumerate(brief.get("project_completion_guidance", []))
    )
    owns_text  = "\n".join(f"- {o}" for o in brief.get("owns", []))
    refs_text  = "\n".join(f"- {r}" for r in brief.get("references", []))

    system = (
        "You are a document engineering assistant generating HSESC management system procedures "
        "for Rio Tinto owner-led projects. Output only valid JSON — no markdown, no code fences."
    )

    style_guide = f"""STYLE REFERENCE — match the example procedure exactly in:
- Section intro length (Level-1: 25-35 words; Level-2: 15-25 words)
- Paragraph length (2-4 sentences; lead-in sentences end with a colon before bullet lists)
- Table format (header row, pre-filled standard rows, [INSERT:] rows for project data)
- INSERT placeholder usage (embedded in sentences, never standalone)
- Orange suggested content (concrete defaults the project may keep or change)
- Impersonal voice (subject is always a role, system, or document — never "you")
- No em-dashes, no contrastive sentence structures, no filler adjectives

EXAMPLE PROCEDURE TEXT:
{example_text}"""

    content_spec = """Content item types for 'content' arrays:
{"type":"paragraph","text":"<text>","suggested":false}  (set suggested:true for orange text)
{"type":"insert","field":"<name>","hint":"<short hint>"}
{"type":"bullet_list","lead":"<lead-in ending with colon>","items":["<text>",{"text":"<orange>","suggested":true},{"insert":true,"field":"<name>"}]}
{"type":"table","caption":"<optional>","suggested":false,"headers":["col"],"rows":[["cell"],["[INSERT: value]"]]}
{"type":"instruction_box","text":"<text>"}"""

    print("  Call 1/2: generating body sections...")
    sections_prompt = f"""Generate body sections for:
PROCEDURE: {brief['procedure_id']} — {brief['title']}
Document: {brief['document_number']} | Element: {brief['management_system_element']}

OWNS: {owns_text}
REFERENCES (cross-reference only, never repeat content): {refs_text}
REQUIREMENTS (every one must appear in the output): {requirements_text}
PROJECT COMPLETION GUIDANCE (embed as [INSERT:] placeholders in the relevant sections): {pcg_text}

{style_guide}
{content_spec}

Return ONLY this JSON:
{{"sections":[{{"title":"<Level-1 section title>","intro":"<25-35 words>","subsections":[{{"title":"<Level-2 or empty string>","intro":"<15-25 words or empty>","level3":[{{"title":"<Level-3>","number":<int>,"content":[<items>]}}],"content":[<items>]}}]}}]}}

RULES:
- Every requirement must appear. Do not omit any.
- Cross-reference owned content in other documents rather than reproducing it.
- Standard PTW phrase: "raise a permit via the site PTW system before commencing work, refer RTPR-HSE-PRO-0021."
- Standard incident phrase: "report all incidents and near misses per project requirements, refer RTPR-HSE-PRO-0009."
- Return only valid JSON."""

    sections_result = _call_claude(client, system, sections_prompt, "sections")

    print("  Call 2/2: generating definitions and references...")
    section_titles = [s.get("title", "") for s in sections_result.get("sections", [])]

    defsrefs_prompt = f"""Generate definitions and references for:
PROCEDURE: {brief['procedure_id']} — {brief['title']}
Body sections: {', '.join(section_titles)}
Referenced documents: {refs_text}

Return ONLY this JSON:
{{"definitions":[{{"term":"<acronym>","definition":"<definition>"}}],"references":[{{"document":"<title>","reference":"<number or standard>","type":"<Internal procedure|External RT|External framework>"}}]}}

Include definitions for every acronym used in the procedure.
Always include RT Risk Management Standard and Appendix C C-1 in references.
Return only valid JSON."""

    defsrefs_result = _call_claude(client, system, defsrefs_prompt, "defs+refs")

    return {
        "sections": sections_result.get("sections", []),
        "definitions": defsrefs_result.get("definitions", []),
        "references": defsrefs_result.get("references", []),
    }


# ── Main ──────────────────────────────────────────────────────────────────────

def read_docx_text(path):
    doc = Document(path)
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and t not in parts:
                    parts.append(t)
    return "\n".join(parts)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--procedure-id",  required=True)
    parser.add_argument("--brief-dir",     required=True)
    parser.add_argument("--template-dir",  required=True)
    parser.add_argument("--output-dir",    required=True)
    parser.add_argument("--inputs-dir",    default="inputs")
    args = parser.parse_args()

    brief_path = os.path.join(args.brief_dir, f"{args.procedure_id}.json")
    if not os.path.exists(brief_path):
        print(f"ERROR: Brief not found: {brief_path}")
        sys.exit(1)

    with open(brief_path, "r", encoding="utf-8") as f:
        brief = json.load(f)

    # Find template
    for candidate in [
        os.path.join(args.inputs_dir, "procedure_template.docx"),
        os.path.join(args.template_dir, "procedure_template.docx"),
    ]:
        if os.path.exists(candidate):
            template_path = candidate
            break
    else:
        print("ERROR: procedure_template.docx not found")
        sys.exit(1)
    print(f"  Template: {template_path}")

    # Find example
    example_path = os.path.join(args.inputs_dir, "example_procedure.docx")
    if os.path.exists(example_path):
        example_text = read_docx_text(example_path)
        if len(example_text) > 10000:
            example_text = example_text[:10000] + "\n[... truncated ...]"
        print(f"  Style reference: example_procedure.docx")
    else:
        example_text = "(No example provided.)"
        print("  WARNING: No example_procedure.docx found.")

    os.makedirs(args.output_dir, exist_ok=True)
    safe = re.sub(r"[^\w\s-]", "", brief["title"].replace(" Procedure","").replace(" Programme",""))
    safe = re.sub(r"\s+", "_", safe.strip())
    output_path = os.path.join(args.output_dir, f"{args.procedure_id}_{safe}.docx")

    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    print(f"Generating: {args.procedure_id} — {brief['title']}")
    generated = generate_content(client, brief, example_text)

    print("Assembling on template...")
    all_elems = []
    all_elems += build_section_0(brief)

    # Section 1 — Purpose and Scope
    all_elems += h1("1  Purpose and Scope")
    all_elems += body(
        "This section defines the purpose of this procedure, the project lifecycle phases within "
        "scope, and the obligations it establishes from project Execution through to Handover."
    )
    all_elems += h2("1.1  Purpose of This Document")
    all_elems += body("This section states the purpose of this procedure and the requirements it establishes for the project.")
    if brief.get("preamble"):
        all_elems += body(brief["preamble"])
    else:
        all_elems += insert_ph("purpose statement", "describe what this procedure governs")

    all_elems += h2("1.2  Scope of This Document")
    all_elems += body("This section defines which activities and lifecycle phases are governed by this document.")
    all_elems += body(
        f"This procedure applies to all Rio Tinto project personnel and all contractor organisations "
        f"performing work under an Owner-Led project. It governs "
        f"{brief['title'].replace(' Procedure','').replace(' Programme','').lower()} "
        f"from project Execution through to Handover."
    )
    all_elems += insert_ph("scope qualifications",
        "note any activities, phases, or locations excluded from this procedure's scope")

    # Section 2 — Roles and Responsibilities
    all_elems += h1("2  Roles and Responsibilities")
    all_elems += body(
        "This section defines the responsibilities for all project roles relevant to this procedure, "
        "covering the Rio Tinto project team and all contractor organisations. Responsibilities are "
        "aligned to governance, oversight, assessment execution, reporting, and control verification."
    )
    all_elems += build_roles_table()

    # Generated content sections
    content_elems, next_n = build_content_sections(generated.get("sections", []))
    all_elems += content_elems

    defs_elems, next_n = build_definitions(next_n, generated.get("definitions", []))
    all_elems += defs_elems

    refs_elems, next_n = build_references(next_n, generated.get("references", []))
    all_elems += refs_elems

    loc_elems, next_n = build_localisation_edits(next_n)
    all_elems += loc_elems

    all_elems += build_traceability(next_n, brief)

    inject_into_template(template_path, output_path, brief, all_elems)
    print(f"Done: {os.path.basename(output_path)}")


if __name__ == "__main__":
    main()
