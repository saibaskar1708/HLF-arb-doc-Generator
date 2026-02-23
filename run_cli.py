"""
run_cli.py  –  Command-line interface for the HLF Reference Letter Generator
─────────────────────────────────────────────────────────────────────────────
Usage:
    python run_cli.py                          # uses defaults below
    python run_cli.py --excel "Lot6.xlsx" --ref-start 262 --lot Lot6
    python run_cli.py --help
"""

import argparse
import os
import sys
from generator import generate

# ─── Defaults ─────────────────────────────────────────────────────────────────
DEFAULT_EXCEL       = "Lot 5 HLF Cases.xlsx"
DEFAULT_SHEET       = "DataFormatted"
DEFAULT_TEMPLATE    = os.path.join("letter_templates", "Reference_Letter_Template.docx")
DEFAULT_OUTPUT_DIR  = "Generated_Reference_Letters"
DEFAULT_REF_START   = 129
DEFAULT_LOT         = "Lot5"
# ──────────────────────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(description="HLF Reference Letter Generator (CLI)")
    parser.add_argument("--excel",     default=DEFAULT_EXCEL,    help="Path to the Excel file")
    parser.add_argument("--sheet",     default=DEFAULT_SHEET,    help="Sheet name in the Excel file")
    parser.add_argument("--template",  default=DEFAULT_TEMPLATE, help="Path to the .docx Jinja2 template")
    parser.add_argument("--output",    default=DEFAULT_OUTPUT_DIR, help="Output folder for generated files")
    parser.add_argument("--ref-start", type=int, default=DEFAULT_REF_START, help="Starting reference number")
    parser.add_argument("--lot",       default=DEFAULT_LOT,      help="Lot label (used in filename)")
    args = parser.parse_args()

    if not os.path.isfile(args.excel):
        print(f"ERROR: Excel file not found: '{args.excel}'")
        sys.exit(1)
    if not os.path.isfile(args.template):
        print(f"ERROR: Template file not found: '{args.template}'")
        sys.exit(1)

    os.makedirs(args.output, exist_ok=True)

    print(f"\n{'─'*55}")
    print(f"  HLF Reference Letter Generator  (CLI)")
    print(f"{'─'*55}")
    print(f"  Excel    : {args.excel}")
    print(f"  Sheet    : {args.sheet}")
    print(f"  Template : {args.template}")
    print(f"  Output   : {args.output}")
    print(f"  Lot      : {args.lot}")
    print(f"  REF from : HLF/SNS/REF/2026/{args.ref_start}")
    print(f"{'─'*55}\n")

    def progress(current, total, msg):
        pct  = int(current / total * 100) if total else 100
        bar  = "█" * (pct // 5) + "░" * (20 - pct // 5)
        print(f"\r  [{bar}] {pct:>3}%  {msg:<50}", end="", flush=True)

    result = generate(
        excel_path        = args.excel,
        template_path     = args.template,
        sheet_name        = args.sheet,
        ref_counter_start = args.ref_start,
        lot_label         = args.lot,
        progress_cb       = progress,
    )
    print()   # newline after progress bar

    # ── Save individual files ──────────────────────────────────────────────
    for fname, buf in result["buffers"]:
        path = os.path.join(args.output, fname)
        buf.seek(0)
        with open(path, "wb") as f:
            f.write(buf.read())

    # ── Save combined doc ──────────────────────────────────────────────────
    combined_path = os.path.join(
        args.output, f"Combined_Reference_Letters_{args.lot}.docx"
    )
    result["combined"].seek(0)
    with open(combined_path, "wb") as f:
        f.write(result["combined"].read())

    # ── Save ZIP ───────────────────────────────────────────────────────────
    zip_path = os.path.join(
        args.output, f"HLF_Reference_Letters_{args.lot}.zip"
    )
    result["zip"].seek(0)
    with open(zip_path, "wb") as f:
        f.write(result["zip"].read())

    # ── Summary ────────────────────────────────────────────────────────────
    print(f"\n{'─'*55}")
    print(f"  ✅  {result['success']} letters generated")
    if result["skipped"]:
        print(f"  ⏭   {result['skipped']} rows skipped (blank Contract No)")
    if result["errors"]:
        print(f"  ❌  {len(result['errors'])} errors:")
        for row, contract, msg in result["errors"]:
            print(f"       Row {row} | {contract} | {msg}")
    print(f"\n  Output folder : {args.output}/")
    print(f"  Combined doc  : {combined_path}")
    print(f"  ZIP           : {zip_path}")
    print(f"{'─'*55}\n")


if __name__ == "__main__":
    main()
