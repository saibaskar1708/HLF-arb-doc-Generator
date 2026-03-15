"""
indiapost_receipts.py — Save one JSON receipt file per article.

Usage:
    python indiapost_receipts.py <input_file.xlsx|csv>

Output:
    Receipts_JSON/<ArticleNumber>.json  (next to the input file)
    Receipts_JSON/receipts_all.json     (all receipts in one file)
"""

import sys
import json
import os

# Windows console encoding fix
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from dotenv import load_dotenv
load_dotenv()

from indiapost_history_pdf import load_ids          # reuse column-detection logic
from indiapost_api import get_receipt_json


def main():
    if len(sys.argv) < 2:
        print("Usage: python indiapost_receipts.py <input_file.xlsx|csv>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"[ERROR] File not found: {input_file}")
        sys.exit(1)

    # Output folder next to the input file
    out_dir = os.path.join(os.path.dirname(os.path.abspath(input_file)), "Receipts_JSON")
    os.makedirs(out_dir, exist_ok=True)

    print(f"[*] Loading tracking IDs from: {input_file}")
    ids = load_ids(input_file)
    print(f"[*] Loaded {len(ids)} tracking IDs")

    if not ids:
        print("[ERROR] No tracking IDs found. Check that your file has a column named "
              "'Tracking ID', 'Article No', 'Article Number', or similar.")
        sys.exit(1)

    print(f"[*] Fetching receipt data from IndiaPost API ...")
    receipts = get_receipt_json(ids)

    # Save individual files
    saved = 0
    not_found = []
    for art_no in ids:
        data = receipts.get(art_no)
        if data:
            path = os.path.join(out_dir, f"{art_no}.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            saved += 1
        else:
            not_found.append(art_no)

    # Save combined file
    combined_path = os.path.join(out_dir, "receipts_all.json")
    with open(combined_path, "w", encoding="utf-8") as f:
        json.dump(receipts, f, indent=2, ensure_ascii=False)

    print(f"\n[OK] Saved {saved} receipt JSON files to: {out_dir}")
    print(f"[OK] Combined file: {combined_path}")

    if not_found:
        print(f"\n[WARN] No data returned for {len(not_found)} article(s):")
        for n in not_found[:20]:
            print(f"       - {n}")
        if len(not_found) > 20:
            print(f"       ... and {len(not_found)-20} more")


if __name__ == "__main__":
    main()
