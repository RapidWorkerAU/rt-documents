"""
get_next_document.py

Determines which document to process next from outputs/documents/,
skipping any that already exist in outputs/documents/fixed/.

Writes to GITHUB_OUTPUT:
  document_name = filename to process now (or DONE)
  next_document  = filename to process after this one (or DONE)
"""

import argparse
import os
import sys


def get_document_list(documents_dir):
    """Get sorted list of .docx files in the documents directory, excluding subdirectories."""
    if not os.path.exists(documents_dir):
        return []
    files = [
        f for f in os.listdir(documents_dir)
        if f.endswith(".docx") and os.path.isfile(os.path.join(documents_dir, f))
    ]
    return sorted(files)


def write_output(key, value):
    github_output = os.environ.get("GITHUB_OUTPUT")
    if github_output:
        with open(github_output, "a") as f:
            f.write(f"{key}={value}\n")
    print(f"  {key} = {value}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--documents-dir", required=True)
    parser.add_argument("--fixed-dir",     required=True)
    parser.add_argument("--start-from",    default="")
    parser.add_argument("--single",        default="")
    args = parser.parse_args()

    all_docs = get_document_list(args.documents_dir)
    if not all_docs:
        print(f"No .docx files found in {args.documents_dir}")
        write_output("document_name", "DONE")
        write_output("next_document",  "DONE")
        return

    print(f"Found {len(all_docs)} documents in {args.documents_dir}")

    # If single document mode
    if args.single:
        if args.single in all_docs:
            write_output("document_name", args.single)
            write_output("next_document",  "DONE")
        else:
            print(f"ERROR: {args.single} not found in {args.documents_dir}")
            write_output("document_name", "DONE")
            write_output("next_document",  "DONE")
        return

    # Determine start position
    if args.start_from and args.start_from in all_docs:
        start_idx = all_docs.index(args.start_from)
    else:
        start_idx = 0

    # Find first document from start_idx that hasn't been fixed yet
    current_doc = None
    current_idx = None

    for i in range(start_idx, len(all_docs)):
        doc = all_docs[i]
        fixed_path = os.path.join(args.fixed_dir, doc)
        if not os.path.exists(fixed_path):
            current_doc = doc
            current_idx = i
            break

    if current_doc is None:
        print("All documents already processed.")
        write_output("document_name", "DONE")
        write_output("next_document",  "DONE")
        return

    write_output("document_name", current_doc)

    # Find next document after current that also needs processing
    next_doc = "DONE"
    for i in range(current_idx + 1, len(all_docs)):
        doc = all_docs[i]
        fixed_path = os.path.join(args.fixed_dir, doc)
        if not os.path.exists(fixed_path):
            next_doc = doc
            break

    write_output("next_document", next_doc)
    print(f"  Will process: {current_doc}")
    print(f"  Next in queue: {next_doc}")


if __name__ == "__main__":
    main()
