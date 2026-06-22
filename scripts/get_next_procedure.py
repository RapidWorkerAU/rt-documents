"""
get_next_procedure.py

Determines the next procedure ID in the registry after the current one.
Writes the result to GITHUB_OUTPUT so the workflow can read it.

Outputs:
  next_id=PRO-XXXX   — next procedure to generate
  next_id=DONE       — all procedures have been generated
"""

import argparse
import json
import os
import sys

PROCEDURE_ORDER = [
    "PRO-0001", "PRO-0002", "PRO-0003", "PRO-0004", "PRO-0005",
    "PRO-0006", "PRO-0007", "PRO-0008", "PRO-0009", "PRO-0010",
    "PRO-0011", "PRO-0012", "PRO-0013", "PRO-0014", "PRO-0015",
    "PRO-0016", "PRO-0017", "PRO-0018", "PRO-0019", "PRO-0020",
    "PRO-0021", "PRO-0022", "PRO-0023", "PRO-0024", "PRO-0025",
    "PRO-0026", "PRO-0027", "PRO-0028", "PRO-0029", "PRO-0030",
    "PRO-0031", "PRO-0032", "PRO-0033", "PRO-0034", "PRO-0035",
    "PRO-0036", "PRO-0037", "PRO-0038", "PRO-0039", "PRO-0040",
    "PRO-0041", "PRO-0042", "PRO-0043", "PRO-0044", "PRO-0045",
    "PRO-0046", "PRO-0047", "PRO-0048", "PRO-0049", "PRO-0050",
    "PRO-0051",
]


def write_output(key: str, value: str):
    """Write a key=value pair to GITHUB_OUTPUT."""
    github_output = os.environ.get("GITHUB_OUTPUT")
    if github_output:
        with open(github_output, "a") as f:
            f.write(f"{key}={value}\n")
    else:
        # Local testing fallback
        print(f"GITHUB_OUTPUT not set. Would write: {key}={value}")
    print(f"Next procedure: {value}")


def main():
    parser = argparse.ArgumentParser(description="Get the next procedure ID in the generation chain")
    parser.add_argument("--current", required=True, help="Current procedure ID, e.g. PRO-0001")
    parser.add_argument("--brief-dir", required=True, help="Directory containing brief JSON files")
    args = parser.parse_args()

    current = args.current.upper().strip()

    if current not in PROCEDURE_ORDER:
        print(f"ERROR: Unknown procedure ID: {current}")
        print(f"Known IDs: {', '.join(PROCEDURE_ORDER)}")
        sys.exit(1)

    current_idx = PROCEDURE_ORDER.index(current)
    next_idx = current_idx + 1

    if next_idx >= len(PROCEDURE_ORDER):
        write_output("next_id", "DONE")
        return

    next_id = PROCEDURE_ORDER[next_idx]

    # Check that the brief exists for the next procedure
    # If extraction failed for next_id, skip it and find the next valid one
    while next_idx < len(PROCEDURE_ORDER):
        candidate = PROCEDURE_ORDER[next_idx]
        brief_path = os.path.join(args.brief_dir, f"{candidate}.json")

        if not os.path.exists(brief_path):
            print(f"  Brief missing for {candidate} — checking next...")
            next_idx += 1
            continue

        # Check if it's a valid (non-failed) brief
        with open(brief_path, "r", encoding="utf-8") as f:
            brief_data = json.load(f)

        if brief_data.get("_extraction_failed"):
            print(f"  Brief for {candidate} has errors — generating anyway with partial data")

        write_output("next_id", candidate)
        return

    # All remaining briefs are missing
    write_output("next_id", "DONE")


if __name__ == "__main__":
    main()
