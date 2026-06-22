"""
generate_procedure.py

Reads a procedure brief JSON and generates a complete .docx procedure document,
following the authoring rules in the HSESC Management System Playbook.

Uses inputs/example_procedure.docx as a live style, structure, and tone reference.
Claude reads the example document and is instructed to match it exactly —
same section depth, same sentence patterns, same table formats, same content density —
but written for the new procedure topic.

Generated document structure:
  Section 0   — How to Use This Document (deletable before issue)
  Section 1   — Purpose and Scope
  Section 2   — Roles and Responsibilities (table)
  Section 3+  — Topic-specific content sections (from brief requirements)
  Section N-3 — Definitions and Acronyms (table)
  Section N-2 — References and Source Information (table)
  Section N-1 — Project Localisation Edits (table, stays in issued document)
  Section N   — Source Requirements Traceability (deletable before issue)
"""

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

import anthropic
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ── Colour constants ──────────────────────────────────────────────────────────
COLOUR_BLACK      = RGBColor(0x2D, 0x2D, 0x2D)
COLOUR_BLUE       = RGBColor(0x00, 0x59, 0xD1)
COLOUR_ORANGE     = RGBColor(0xC5, 0x5A, 0x11)
COLOUR_GREEN      = RGBColor(0x37, 0x56, 0x23)
COLOUR_HEADING    = RGBColor(0x1F, 0x38, 0x64)
COLOUR_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)


# ── Document helpers ──────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_colour: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_colour)
    tcPr.append(shd)


def para_border_bottom(para, colour="2E75B6", size=6):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), colour)
    pBdr.append(bottom)
    pPr.append(pBdr)


def run(para, text, bold=False, colour=None, size_pt=10, italic=False):
    r = para.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size_pt)
    r.font.name = "Arial"
    if colour:
        r.font.color.rgb = colour
    return r


def h1(doc, number, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    para_border_bottom(p)
    run(p, f"{number}  {title}", bold=True, colour=COLOUR_HEADING, size_pt=13)
    return p


def h2(doc, number, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run(p, f"{number}  {title}", bold=True, colour=COLOUR_HEADING, size_pt=11)
    return p


def h3(doc, number, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    run(p, f"{number}  {title}", bold=True, colour=COLOUR_BLACK, size_pt=10)
    return p


def body(doc, text, colour=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    run(p, text, colour=colour or COLOUR_BLACK, size_pt=10)
    return p


def bullet(doc, text, colour=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run(p, text, colour=colour or COLOUR_BLACK, size_pt=10)
    return p


def insert_ph(doc, field, hint=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    text = f"[INSERT: {field}]"
    if hint:
        text += f" — {hint}"
    run(p, text, bold=True, colour=COLOUR_BLUE, size_pt=10)
    return p


def shaded_box(doc, text, bg, text_colour=None, bold=False):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, bg)
    cell.paragraphs[0].clear()
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run(p, text, bold=bold, colour=text_colour or COLOUR_BLACK, size_pt=10)
    doc.add_paragraph()
    return tbl


def make_table(doc, rows, cols, widths_cm):
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.style = "Table Grid"
    for row in tbl.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])
    return tbl


def hdr_row(tbl, row_idx, labels, bg="1F3864"):
    row = tbl.rows[row_idx]
    for i, label in enumerate(labels):
        if i < len(row.cells):
            set_cell_bg(row.cells[i], bg)
            p = row.cells[i].paragraphs[0]
            run(p, label, bold=True, colour=COLOUR_WHITE, size_pt=9)


def read_docx_text(path: str) -> str:
    """Extract all paragraph and table text from a docx."""
    doc = Document(path)
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and t not in parts:
                    parts.append(t)
    return "\n".join(parts)


# ── Section 0 ─────────────────────────────────────────────────────────────────

def build_section_0(doc, brief):
    shaded_box(
        doc,
        "INSTRUCTION: Read, action and delete this box before this procedure is issued. "
        "Delete the entire Section 0 before this document is issued for review or approval. "
        "It must not appear in any approved version of this procedure.",
        bg="FFF2CC", text_colour=COLOUR_ORANGE, bold=True
    )

    h1(doc, "0", "How to Use This Document")
    body(doc,
        "This section explains the purpose of the template, how it is structured, the sequence for completing it "
        "and the checks required before it is issued. It must be read in full before any section of this document "
        "is completed. Delete Section 0 in its entirety before issue."
    )

    h2(doc, "0.1", "Purpose of This Template")
    body(doc,
        f"This template supports the preparation of a project-specific {brief['title']} for projects delivered "
        "under an Owner-Led model. Under this model, Rio Tinto directly manages all project activities, and this "
        "procedure applies to all personnel on the project, including Rio Tinto staff and all contractor organisations."
    )
    body(doc,
        "This template is provided as a 90% draft. Fixed content reflects the minimum Rio Tinto requirement and "
        "must not be reduced. The remaining 10% covers role names, register locations, schedules and project-specific "
        "triggers, and must be completed by the project team before this document is issued."
    )

    h2(doc, "0.2", "How This Document is Structured")
    body(doc, "This document contains four content types:")

    p1 = doc.add_paragraph()
    run(p1, "Fixed content ", bold=True, colour=COLOUR_BLACK, size_pt=10)
    run(p1, "is plain black text reflecting minimum Rio Tinto requirements. Fixed content must not be reduced, deleted or overridden.", colour=COLOUR_BLACK, size_pt=10)

    p2 = doc.add_paragraph()
    run(p2, "[INSERT: placeholder text] ", bold=True, colour=COLOUR_BLUE, size_pt=10)
    run(p2, "appears in blue bold. Each placeholder must be replaced with accurate, project-specific information before issue.", colour=COLOUR_BLACK, size_pt=10)

    p3 = doc.add_paragraph()
    run(p3, "Orange text ", colour=COLOUR_ORANGE, size_pt=10)
    run(p3, "marks specific who, what, when, or how detail included as an industry best practice suggestion. The project team may retain, modify, or replace orange detail to suit their own organisational context.", colour=COLOUR_BLACK, size_pt=10)

    h2(doc, "0.3", "Pre-Issue Checklist")
    body(doc,
        "Before submitting this document for review or approval, confirm all items below are complete, "
        "then delete this checklist together with the rest of Section 0."
    )

    checklist = [
        "Cover page is complete including document number, revision, date and signatories",
        "All [INSERT:] placeholder text has been replaced with project-specific information",
        "All instruction boxes have been read, actioned and deleted",
        "All role titles match the actual project organisational structure",
        "All system and register locations are confirmed and inserted",
        "All workshop and meeting dates are confirmed against the project programme",
        "Related procedures listed in the References section are issued or being issued concurrently",
        "Section 17 (Source Requirements Traceability) has been completed and will be deleted before final issue",
        "Section 16 (Project Localisation Edits) has been populated with all changes made during preparation",
        "Section 0 has been deleted in its entirety",
    ]

    tbl = make_table(doc, rows=len(checklist) + 1, cols=2, widths_cm=[14, 2])
    hdr_row(tbl, 0, ["Check", "Complete?"])
    for i, item in enumerate(checklist):
        run(tbl.rows[i+1].cells[0].paragraphs[0], item, colour=COLOUR_BLACK, size_pt=9)
        run(tbl.rows[i+1].cells[1].paragraphs[0], "☐", colour=COLOUR_BLACK, size_pt=9)
    doc.add_paragraph()


# ── Cover ─────────────────────────────────────────────────────────────────────

def build_cover(doc, brief):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, "RioProjects HSESC", bold=True, colour=COLOUR_HEADING, size_pt=10)

    doc.add_paragraph()
    pt = doc.add_paragraph()
    pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(pt, brief["title"], bold=True, colour=COLOUR_HEADING, size_pt=18)

    doc.add_paragraph()
    pi = doc.add_paragraph()
    pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(pi, "[Enter Project Name]", bold=True, colour=COLOUR_BLUE, size_pt=12)

    doc.add_paragraph()

    # Document control
    tbl = make_table(doc, rows=2, cols=3, widths_cm=[6, 4, 6])
    hdr_row(tbl, 0, ["Document Number", "Revision", "Date"])
    run(tbl.rows[1].cells[0].paragraphs[0], brief["document_number"], size_pt=9)
    run(tbl.rows[1].cells[1].paragraphs[0], "Rev 0", size_pt=9)
    run(tbl.rows[1].cells[2].paragraphs[0], "[INSERT: date]", bold=True, colour=COLOUR_BLUE, size_pt=9)
    doc.add_paragraph()

    # Signatories
    sig = make_table(doc, rows=5, cols=4, widths_cm=[3, 5, 5, 3])
    hdr_row(sig, 0, [" Status", "Name", "Position", "Signature"])
    for i, (status, name, pos) in enumerate([
        ("Originator", "[INSERT: name]", "[INSERT: position and organisation]"),
        ("Reviewed",   "[INSERT: name]", "[INSERT: position]"),
        ("Reviewed",   "[INSERT: name]", "[INSERT: position]"),
        ("Approved",   "[INSERT: name]", "[INSERT: position]"),
    ]):
        run(sig.rows[i+1].cells[0].paragraphs[0], status, size_pt=9)
        run(sig.rows[i+1].cells[1].paragraphs[0], name, bold=True, colour=COLOUR_BLUE, size_pt=9)
        run(sig.rows[i+1].cells[2].paragraphs[0], pos, bold=True, colour=COLOUR_BLUE, size_pt=9)
    doc.add_paragraph()


# ── Section builders ──────────────────────────────────────────────────────────

def build_purpose_scope(doc, brief, n):
    h1(doc, n, "Purpose and Scope")
    body(doc, brief.get("preamble") or
        f"This section defines the purpose of this procedure and the project activities within scope.")

    h2(doc, f"{n}.1", "Purpose of This Document")
    body(doc,
        f"This section states the purpose of this procedure and the requirements it establishes for the project "
        f"from Execution through to Handover."
    )
    body(doc,
        f"This procedure meets the following purpose requirements:"
    )
    if brief.get("preamble"):
        body(doc, brief["preamble"])
    else:
        insert_ph(doc, "purpose statement",
            "describe what this procedure governs and the project context it applies to")

    h2(doc, f"{n}.2", "Scope of This Document")
    body(doc,
        f"This section defines which activities and lifecycle phases are governed by this document."
    )
    body(doc,
        f"This procedure applies to all Rio Tinto project personnel and all contractor organisations "
        f"performing work under an Owner-Led project. It governs the requirements for "
        f"{brief['title'].replace(' Procedure','').replace(' Programme','').lower()} "
        f"from project Execution through to Handover."
    )
    insert_ph(doc, "scope qualifications",
        "note any activities, phases, or locations excluded from this procedure's scope")
    return n + 1


def build_roles(doc, brief, n):
    h1(doc, n, "Roles and Responsibilities")
    body(doc,
        "This section defines the responsibilities for all project roles relevant to this procedure, "
        "covering the Rio Tinto project team and all contractor organisations. Responsibilities are "
        "aligned to governance, oversight, assessment execution, reporting, and control verification."
    )

    tbl = make_table(doc, rows=1, cols=4, widths_cm=[8, 1.5, 1.5, 5])
    hdr_row(tbl, 0, ["Responsibilities", "Company", "Contractor", "Section"])

    role_groups = [
        ("Project Management", [
            ("Holds overall accountability for the project compliance with this procedure and approves this procedure and any material amendment.", "✓", ""),
        ]),
        ("HSE/CSP Management", [
            ("Maintains and implements this procedure and confirms all activities meet its requirements before work proceeds.", "✓", "✓"),
            ("Reviews and approves all required assessments and plans produced under this procedure.", "✓", "✓"),
            ("Reports performance against this procedure at the agreed project HSESC performance meeting.", "✓", "✓"),
        ]),
        ("Frontline Supervision", [
            ("Ensures work crew compliance with this procedure before and during all relevant task executions.", "✓", "✓"),
            ("Stops work when the requirements of this procedure cannot be met and notifies the HSESC Manager.", "✓", "✓"),
        ]),
        ("All Project Personnel / Workers", [
            ("Complies with the requirements of this procedure in all relevant work activities.", "✓", "✓"),
            ("Reports any non-compliance with this procedure immediately to their supervisor.", "✓", "✓"),
        ]),
    ]

    for role_name, responsibilities in role_groups:
        # Role group header
        grp_row = tbl.add_row()
        set_cell_bg(grp_row.cells[0], "D6E4F0")
        merged = grp_row.cells[0].merge(grp_row.cells[3])
        run(merged.paragraphs[0], role_name, bold=True, colour=COLOUR_HEADING, size_pt=9)

        for resp_text, company, contractor in responsibilities:
            r = tbl.add_row()
            run(r.cells[0].paragraphs[0], resp_text, colour=COLOUR_BLACK, size_pt=9)
            run(r.cells[1].paragraphs[0], company, bold=True, colour=COLOUR_BLACK, size_pt=9)
            run(r.cells[2].paragraphs[0], contractor, bold=True, colour=COLOUR_BLACK, size_pt=9)

        ins = tbl.add_row()
        run(ins.cells[0].paragraphs[0], "[INSERT: add further responsibilities as required]",
            bold=True, colour=COLOUR_BLUE, size_pt=9)

    doc.add_paragraph()
    return n + 1


def build_content_sections(doc, n, sections_data):
    """
    Build all generated topic sections.
    sections_data is a list of section dicts from Claude's output.
    """
    for sec in sections_data:
        title = sec.get("title", "")
        intro = sec.get("intro", "")
        subsections = sec.get("subsections", [])

        h1(doc, n, title)
        if intro:
            body(doc, intro)

        sub_n = 1
        for sub in subsections:
            sub_title = sub.get("title", "")
            sub_intro = sub.get("intro", "")

            if sub_title:
                h2(doc, f"{n}.{sub_n}", sub_title)

            if sub_intro:
                body(doc, sub_intro)

            # Level-3 subsections
            for l3 in sub.get("level3", []):
                l3_title = l3.get("title", "")
                if l3_title:
                    h3(doc, f"{n}.{sub_n}.{l3.get('number', 1)}", l3_title)
                for item in l3.get("content", []):
                    _render_content_item(doc, item)

            # Content items at level-2 (no level-3)
            for item in sub.get("content", []):
                _render_content_item(doc, item)

            if sub_title:
                sub_n += 1

        n += 1
    return n


def _render_content_item(doc, item):
    """Render a single content item. item is a dict with a 'type' key."""
    t = item.get("type", "paragraph")

    if t == "paragraph":
        text = item.get("text", "")
        colour = COLOUR_ORANGE if item.get("suggested") else COLOUR_BLACK
        body(doc, text, colour=colour)

    elif t == "insert":
        insert_ph(doc, item.get("field", "value"), item.get("hint", ""))

    elif t == "bullet_list":
        lead = item.get("lead", "")
        if lead:
            body(doc, lead)
        for b in item.get("items", []):
            if isinstance(b, dict) and b.get("insert"):
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.left_indent = Cm(1.0)
                run(p, f"[INSERT: {b.get('field', 'value')}]",
                    bold=True, colour=COLOUR_BLUE, size_pt=10)
            else:
                colour = COLOUR_ORANGE if (isinstance(b, dict) and b.get("suggested")) else COLOUR_BLACK
                text = b.get("text", b) if isinstance(b, dict) else b
                bullet(doc, str(text), colour=colour)

    elif t == "table":
        headers = item.get("headers", [])
        rows_data = item.get("rows", [])
        caption = item.get("caption", "")
        if caption:
            body(doc, caption)
        if headers:
            col_w = [16.0 / len(headers)] * len(headers)
            tbl = make_table(doc, rows=len(rows_data) + 1, cols=len(headers), widths_cm=col_w)
            hdr_row(tbl, 0, headers)
            for ri, row_vals in enumerate(rows_data):
                for ci, val in enumerate(row_vals):
                    if ci < len(tbl.rows[ri+1].cells):
                        val_str = str(val)
                        if val_str.startswith("[INSERT:"):
                            run(tbl.rows[ri+1].cells[ci].paragraphs[0],
                                val_str, bold=True, colour=COLOUR_BLUE, size_pt=9)
                        elif item.get("suggested"):
                            run(tbl.rows[ri+1].cells[ci].paragraphs[0],
                                val_str, colour=COLOUR_ORANGE, size_pt=9)
                        else:
                            run(tbl.rows[ri+1].cells[ci].paragraphs[0],
                                val_str, colour=COLOUR_BLACK, size_pt=9)
            doc.add_paragraph()

    elif t == "instruction_box":
        shaded_box(doc, item.get("text", ""), bg="FFF2CC",
                   text_colour=COLOUR_ORANGE, bold=True)


def build_definitions(doc, n, definitions):
    h1(doc, n, "Definitions and Acronyms")
    body(doc,
        "The following terms are used throughout this procedure. Where a term is also used in other "
        "project HSESC procedures, the definition below applies consistently across the system."
    )
    tbl = make_table(doc, rows=1, cols=2, widths_cm=[4, 12])
    hdr_row(tbl, 0, ["Term", "Definition"])
    for d in definitions:
        r = tbl.add_row()
        run(r.cells[0].paragraphs[0], d.get("term", ""), bold=True, colour=COLOUR_BLACK, size_pt=9)
        run(r.cells[1].paragraphs[0], d.get("definition", ""), colour=COLOUR_BLACK, size_pt=9)
    ins = tbl.add_row()
    run(ins.cells[0].paragraphs[0], "[INSERT: add project-specific terms]",
        bold=True, colour=COLOUR_BLUE, size_pt=9)
    run(ins.cells[1].paragraphs[0], "[INSERT: definition]",
        bold=True, colour=COLOUR_BLUE, size_pt=9)
    doc.add_paragraph()
    return n + 1


def build_references(doc, n, references):
    h1(doc, n, "References and Source Information")
    body(doc,
        "This section lists the internal and external documents this procedure relies on. "
        "This section remains in the procedure once published. Where a referenced document is updated, "
        "this procedure is reviewed to confirm the reference remains current."
    )
    tbl = make_table(doc, rows=1, cols=3, widths_cm=[8, 4, 4])
    hdr_row(tbl, 0, ["Document", "Reference", "Type"])
    for ref in references:
        r = tbl.add_row()
        run(r.cells[0].paragraphs[0], ref.get("document", ""), colour=COLOUR_BLACK, size_pt=9)
        run(r.cells[1].paragraphs[0], ref.get("reference", ""), colour=COLOUR_BLACK, size_pt=9)
        run(r.cells[2].paragraphs[0], ref.get("type", ""), colour=COLOUR_BLACK, size_pt=9)
    ins = tbl.add_row()
    run(ins.cells[0].paragraphs[0], "[INSERT: further referenced document, where applicable]",
        bold=True, colour=COLOUR_BLUE, size_pt=9)
    run(ins.cells[1].paragraphs[0], "[INSERT: reference]", bold=True, colour=COLOUR_BLUE, size_pt=9)
    run(ins.cells[2].paragraphs[0], "[INSERT: type]", bold=True, colour=COLOUR_BLUE, size_pt=9)
    doc.add_paragraph()
    return n + 1


def build_localisation_edits(doc, n):
    h1(doc, n, "Project Localisation Edits")
    body(doc,
        "This section captures the project-specific edits made to this document as it is prepared for "
        "project-specific use. This section stays in the document once published and is updated at each "
        "document revision. Record all additions, deletions, and wording changes from the original template."
    )

    def change_table(heading_text):
        p = doc.add_paragraph()
        run(p, heading_text, bold=True, colour=COLOUR_BLACK, size_pt=10)
        tbl = make_table(doc, rows=4, cols=2, widths_cm=[13, 3])
        hdr_row(tbl, 0, ["Description of the Change", "Section"])
        for i in range(1, 4):
            run(tbl.rows[i].cells[0].paragraphs[0], "[INSERT: description]",
                bold=True, colour=COLOUR_BLUE, size_pt=9)
            run(tbl.rows[i].cells[1].paragraphs[0], "[INSERT: section]",
                bold=True, colour=COLOUR_BLUE, size_pt=9)
        doc.add_paragraph()

    change_table("Records of wording changes")
    change_table("Records of new content additions")
    change_table("Records of content removal")
    return n + 1


def build_traceability(doc, n, brief):
    shaded_box(
        doc,
        "REMOVABLE SECTION — Source Requirements Traceability. "
        "Delete this entire section before the document is issued for approval. "
        "Table A records requirements addressed in this procedure. "
        "Table B records requirements not fully addressed, with written justification for each.",
        bg="FCE4D6", text_colour=COLOUR_ORANGE, bold=True
    )

    h1(doc, n, "Source Requirements Traceability (DELETE BEFORE ISSUE)")

    # Table A
    h2(doc, f"{n}.1", "Requirements Addressed in This Procedure")
    tbl_a = make_table(doc, rows=1, cols=4, widths_cm=[1.5, 2.5, 6, 6])
    hdr_row(tbl_a, 0, ["Section", "Source", "Requirement", "Status"])
    for src in brief.get("external_sources", []):
        r = tbl_a.add_row()
        run(r.cells[0].paragraphs[0], "[INSERT: section]",
            bold=True, colour=COLOUR_BLUE, size_pt=9)
        run(r.cells[1].paragraphs[0],
            f"{src.get('framework','')} {src.get('clause','')}".strip(),
            colour=COLOUR_BLACK, size_pt=9)
        run(r.cells[2].paragraphs[0], src.get("summary", ""),
            colour=COLOUR_BLACK, size_pt=9)
        run(r.cells[3].paragraphs[0], "Full", colour=COLOUR_BLACK, size_pt=9)
    doc.add_paragraph()

    # Table B
    h2(doc, f"{n}.2", "Requirements Not Fully Addressed in This Procedure")
    body(doc,
        "List any requirements from the external sources above that are not fully addressed in this "
        "document, with a specific written justification for each. Acceptable justifications: "
        "(a) the requirement is owned by a named standalone procedure and this document cross-references it; "
        "(b) excluded by a specific business decision endorsed by project leadership, named here; "
        "or (c) addressed through an operational control outside this document, identified here."
    )
    tbl_b = make_table(doc, rows=2, cols=3, widths_cm=[6, 3, 7])
    hdr_row(tbl_b, 0, ["Requirement Not Covered", "Source", "Justification"])
    ins = tbl_b.rows[1]
    run(ins.cells[0].paragraphs[0], "[INSERT: description of requirement not covered]",
        bold=True, colour=COLOUR_BLUE, size_pt=9)
    run(ins.cells[1].paragraphs[0], "[INSERT: source and clause]",
        bold=True, colour=COLOUR_BLUE, size_pt=9)
    run(ins.cells[2].paragraphs[0], "[INSERT: specific written justification]",
        bold=True, colour=COLOUR_BLUE, size_pt=9)
    doc.add_paragraph()
    return n + 1


# ── Claude generation ─────────────────────────────────────────────────────────

def _call_claude(client, system, prompt, label=""):
    """Make a single Claude API call with retry logic. Returns parsed JSON."""
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
    """
    Generate procedure body content across two API calls to avoid token limits.
    Call 1: body sections (the large content).
    Call 2: definitions and references (small, fast).
    Results are merged into one dict before document assembly.
    """

    requirements_text = "\n".join(
        f"- [{r.get('tag','RT Standard')}] under heading '{r.get('section_heading','')}': {r.get('text','')}"
        for r in brief.get("requirements", [])
    )
    pcg_text = "\n".join(
        f"{i+1}. {item}"
        for i, item in enumerate(brief.get("project_completion_guidance", []))
    )
    owns_text = "\n".join(f"- {o}" for o in brief.get("owns", []))
    refs_text = "\n".join(f"- {r}" for r in brief.get("references", []))

    system = (
        "You are a document engineering assistant generating HSESC management system procedures "
        "for Rio Tinto owner-led projects. You produce technically precise, production-ready "
        "procedure content. Output only valid JSON — no markdown, no code fences, no commentary."
    )

    style_block = f"""════════════════════════════════════════════════════════
STYLE REFERENCE — EXAMPLE PROCEDURE
════════════════════════════════════════════════════════
The following is the full text of a completed procedure from the same management system.
Match its structure, heading depth, sentence patterns, paragraph length, table formats,
INSERT placeholder usage, and content density exactly — but written for the new topic.

Key patterns to replicate:
1. Each Level-1 section starts with a 25-35 word intro stating what the section covers.
2. Each Level-2 section starts with a 15-25 word intro stating what the subsection defines.
3. Paragraphs are 2-4 sentences. No single-sentence paragraphs except lead-ins.
4. Tables are pre-filled: fixed rows for standard content, [INSERT:] rows for project data.
5. INSERT placeholders are embedded inside sentences — never as standalone lines.
6. Orange suggested content provides a concrete default the project may keep or change.
7. Bullet lists always have a lead-in sentence ending with a colon.
8. Impersonal voice throughout — subject is always a role, system, or document.

EXAMPLE PROCEDURE TEXT:
{example_text}
════════════════════════════════════════════════════════"""

    content_item_spec = """Content item types for 'content' arrays:

Paragraph:
  {"type": "paragraph", "text": "<text>", "suggested": false}
  Set suggested: true for orange best-practice text the project may modify.

INSERT placeholder:
  {"type": "insert", "field": "<field name>", "hint": "<short hint>"}

Bullet list:
  {"type": "bullet_list", "lead": "<lead-in sentence ending with colon>", "items": [
    "<bullet text>",
    {"text": "<suggested bullet>", "suggested": true},
    {"insert": true, "field": "<field name>"}
  ]}

Table:
  {"type": "table", "caption": "<optional caption>", "suggested": false,
    "headers": ["<col>", "<col>"],
    "rows": [["<cell>", "<cell>"], ["[INSERT: value]", "[INSERT: value]"]]
  }

Instruction box (use sparingly):
  {"type": "instruction_box", "text": "<instruction text>"}"""

    # ── Call 1: body sections ─────────────────────────────────────────────────
    print("  Call 1/2: generating body sections...")
    sections_prompt = f"""Generate the body sections for this procedure:

PROCEDURE: {brief['procedure_id']} — {brief['title']}
Document number: {brief['document_number']}
Management system element: {brief['management_system_element']}

THIS DOCUMENT OWNS:
{owns_text}

THIS DOCUMENT REFERENCES (cross-reference only — do not repeat content):
{refs_text}

REQUIREMENTS TO ADDRESS (every one must appear in the output):
{requirements_text}

PROJECT COMPLETION GUIDANCE (embed as [INSERT:] placeholders in the relevant sections):
{pcg_text}

{style_block}

{content_item_spec}

Return ONLY this JSON — nothing else:

{{
  "sections": [
    {{
      "title": "<Level-1 section title>",
      "intro": "<25-35 word introduction>",
      "subsections": [
        {{
          "title": "<Level-2 title, or empty string>",
          "intro": "<15-25 word introduction, or empty string>",
          "level3": [
            {{
              "title": "<Level-3 title>",
              "number": <integer>,
              "content": [ <content items> ]
            }}
          ],
          "content": [ <content items> ]
        }}
      ]
    }}
  ]
}}

RULES:
- Every requirement must be addressed. Do not omit any.
- Do not reproduce content owned by referenced documents — cross-reference instead.
- Standard PTW cross-reference: "raise a permit via the site PTW system before commencing work, refer RTPR-HSE-PRO-0021."
- Standard incident cross-reference: "report all incidents and near misses per the project incident classification and notification requirements, refer RTPR-HSE-PRO-0009."
- Return only valid JSON."""

    sections_result = _call_claude(client, system, sections_prompt, label="sections")

    # ── Call 2: definitions and references ────────────────────────────────────
    print("  Call 2/2: generating definitions and references...")

    # Summarise what sections were generated so Claude can produce matching definitions
    section_titles = [s.get("title", "") for s in sections_result.get("sections", [])]
    section_summary = ", ".join(section_titles)

    defsrefs_prompt = f"""Generate the definitions and references for this procedure.

PROCEDURE: {brief['procedure_id']} — {brief['title']}

The procedure body covers these sections: {section_summary}

It references these documents:
{refs_text}

Return ONLY this JSON — nothing else:

{{
  "definitions": [
    {{"term": "<acronym or term used in the procedure>", "definition": "<clear definition>"}}
  ],
  "references": [
    {{"document": "<full document title>", "reference": "<document number or standard>", "type": "<Internal procedure | External RT | External framework>"}}
  ]
}}

RULES:
- definitions must cover every acronym and technical term used in the procedure sections above.
- Always include: HSESC, RT, ALARP, JHA, PTW, SWMS, MoC, PCG, PPE where relevant to the topic.
- references must include every internal procedure cross-referenced in the body sections.
- Always include as references: RT Risk Management Standard, Appendix C C-1, and any IFC/ICMM standards relevant to the topic.
- Return only valid JSON."""

    defsrefs_result = _call_claude(client, system, defsrefs_prompt, label="defs+refs")

    # ── Merge results ─────────────────────────────────────────────────────────
    return {
        "sections": sections_result.get("sections", []),
        "definitions": defsrefs_result.get("definitions", []),
        "references": defsrefs_result.get("references", []),
    }


# ── Document assembly ─────────────────────────────────────────────────────────

def assemble_document(brief, generated, output_path):
    doc = Document()

    # Page setup — A4
    sec = doc.sections[0]
    sec.page_height = Cm(29.7)
    sec.page_width = Cm(21.0)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, attr, Cm(2.54))

    build_cover(doc, brief)
    doc.add_page_break()

    build_section_0(doc, brief)
    doc.add_page_break()

    # TOC placeholder
    p = doc.add_paragraph()
    run(p, "Contents", bold=True, colour=COLOUR_HEADING, size_pt=13)
    para_border_bottom(p)
    body(doc,
        "[Right-click here and select 'Update Field' to generate the table of contents "
        "after all content is finalised and placeholders are completed.]"
    )
    doc.add_paragraph()
    doc.add_page_break()

    # Body
    n = 1
    n = build_purpose_scope(doc, brief, n)
    n = build_roles(doc, brief, n)
    n = build_content_sections(doc, n, generated.get("sections", []))
    n = build_definitions(doc, n, generated.get("definitions", []))
    n = build_references(doc, n, generated.get("references", []))
    n = build_localisation_edits(doc, n)

    doc.add_page_break()
    build_traceability(doc, n, brief)

    doc.save(output_path)
    print(f"  Saved: {output_path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--procedure-id",  required=True)
    parser.add_argument("--brief-dir",     required=True)
    parser.add_argument("--template-dir",  required=True)
    parser.add_argument("--output-dir",    required=True)
    parser.add_argument("--inputs-dir",    default="inputs",
                        help="Directory containing example_procedure.docx")
    args = parser.parse_args()

    brief_path = os.path.join(args.brief_dir, f"{args.procedure_id}.json")
    if not os.path.exists(brief_path):
        print(f"ERROR: Brief not found: {brief_path}")
        sys.exit(1)

    with open(brief_path, "r", encoding="utf-8") as f:
        brief = json.load(f)

    # Load example procedure for style reference
    example_path = os.path.join(args.inputs_dir, "example_procedure.docx")
    if not os.path.exists(example_path):
        # Fall back to any .docx in inputs that isn't the playbook
        for fname in os.listdir(args.inputs_dir):
            if fname.endswith(".docx") and "playbook" not in fname.lower():
                example_path = os.path.join(args.inputs_dir, fname)
                print(f"  Using fallback example: {fname}")
                break

    if os.path.exists(example_path):
        print(f"  Loading style reference: {os.path.basename(example_path)}")
        example_text = read_docx_text(example_path)
        # Truncate if very long — keep enough for style reference without blowing context
        if len(example_text) > 12000:
            example_text = example_text[:12000] + "\n[... remainder of example truncated for context ...]"
    else:
        print("  WARNING: No example_procedure.docx found in inputs/. Generating without style reference.")
        example_text = "(No example procedure provided.)"

    os.makedirs(args.output_dir, exist_ok=True)

    safe = re.sub(r"[^\w\s-]", "", brief["title"].replace(" Procedure","").replace(" Programme",""))
    safe = re.sub(r"\s+", "_", safe.strip())
    output_path = os.path.join(args.output_dir, f"{args.procedure_id}_{safe}.docx")

    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    print(f"Generating content for {args.procedure_id}: {brief['title']}...")
    generated = generate_content(client, brief, example_text)

    print("Building document...")
    assemble_document(brief, generated, output_path)

    print(f"Done: {os.path.basename(output_path)}")


if __name__ == "__main__":
    main()
