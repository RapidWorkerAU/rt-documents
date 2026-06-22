# rt-documents

Automated HSESC Management System document generation pipeline for Rio Tinto owner-led projects.

## What this does

This pipeline reads the HSESC Management System Playbook and automatically generates all 51 project procedure documents as Word (.docx) files, following the authoring rules and content requirements defined in the playbook.

Each generated document includes:
- Section 0 — How to Use (pre-issue instruction block, deletable before issue)
- Section 1 — Purpose and Scope
- Section 2 — Roles and Responsibilities (pre-filled table)
- Topic-specific content sections (generated from the procedure brief)
- Definitions and Acronyms (pre-filled table)
- References and Source Information (pre-filled table)
- Project Localisation Edits (standing change log, stays in issued document)
- Section 17 — Source Requirements Traceability with Table A and Table B (deletable before issue)

---

## Folder structure

```
rt-documents/
├── .github/
│   └── workflows/
│       ├── orchestrate.yml          # Step 1: extract briefs and start the chain
│       └── generate_procedure.yml   # Step 2: generate one procedure, trigger next
├── briefs/                          # Auto-generated JSON briefs (one per procedure)
├── inputs/
│   ├── playbook.docx                # Your HSESC MS Playbook v8 (source of truth)
│   └── procedure_template.docx      # Blank procedure Word template
├── outputs/                         # Generated .docx procedures committed here
├── scripts/
│   ├── extract_briefs.py            # Parses playbook → 51 JSON briefs
│   ├── generate_procedure.py        # Calls Claude API → builds .docx
│   └── get_next_procedure.py        # Chain logic: determines next procedure to run
└── template/                        # Reserved for future template assets
```

---

## Setup

### 1. Add your Anthropic API key to GitHub Secrets

1. Go to your repository on GitHub
2. Settings → Secrets and variables → Actions
3. Click **New repository secret**
4. Name: `ANTHROPIC_API_KEY`
5. Value: your Anthropic API key (starts with `sk-ant-...`)

### 2. Enable workflow permissions

1. Go to Settings → Actions → General
2. Under **Workflow permissions**, select **Read and write permissions**
3. Check **Allow GitHub Actions to create and approve pull requests**
4. Click Save

### 3. Confirm inputs are in place

Make sure the following files exist in the `inputs/` folder:
- `playbook.docx` — the HSESC Management System Playbook (v8 or later)
- `procedure_template.docx` — the blank procedure Word template

---

## Running the pipeline

### Full run (all 51 procedures)

1. Go to your repository on GitHub
2. Actions → **01 - Extract Briefs and Start Generation**
3. Click **Run workflow**
4. Leave **Start from** as `PRO-0001`
5. Leave **Dry run** as `false`
6. Click **Run workflow**

The pipeline will:
1. Extract all 51 procedure briefs from the playbook (takes ~5 minutes)
2. Generate PRO-0001 and commit it to `outputs/`
3. Automatically trigger PRO-0002, then PRO-0003, and so on through PRO-0051
4. Each procedure takes approximately 2-4 minutes to generate

Total runtime for all 51 procedures: approximately 2-4 hours.

### Dry run (extract briefs only, no document generation)

Set **Dry run** to `true`. This extracts and commits all 51 JSON briefs to the `briefs/` folder without generating any Word documents. Useful for checking the brief extraction before committing to a full run.

### Resume from a specific procedure

If the pipeline fails partway through (e.g. at PRO-0023), you can resume:
1. Actions → **01 - Extract Briefs and Start Generation**
2. Set **Start from** to `PRO-0023`
3. Run — it will skip the brief extraction (briefs already exist) and start generating from PRO-0023

Or trigger a single procedure directly:
1. Actions → **02 - Generate Procedure Document**
2. Set **Procedure ID** to the one you want, e.g. `PRO-0023`
3. Run — generates that procedure and then continues chaining to the next one

---

## Output

Generated documents are committed to the `outputs/` folder with filenames like:
```
PRO-0001_Legal_and_Other_Requirements.docx
PRO-0002_HSESC_Risk_Management.docx
...
PRO-0051_Explosives_Management.docx
```

---

## After generation

Each generated document is a 90% draft. Before issuing any procedure:
1. Open the document
2. Read Section 0 in full
3. Complete all `[INSERT: ...]` placeholders
4. Review orange suggested content and retain, modify, or replace as needed
5. Complete Section 17 (Source Requirements Traceability) during review
6. Delete Section 0 and Section 17 before issuing for approval
7. Record all changes in Section 16 (Project Localisation Edits)

---

## Troubleshooting

**Brief extraction fails for a procedure**
The script writes a placeholder brief so the pipeline can continue. The generated document will have minimal content. Re-run just that procedure after fixing the issue.

**Workflow does not trigger the next procedure**
Check that **Workflow permissions** are set to read/write in repository settings (see Setup step 2).

**API rate limit errors**
The scripts include retry logic and courtesy pauses between calls. If you hit rate limits, wait 10 minutes and resume from the failed procedure.
