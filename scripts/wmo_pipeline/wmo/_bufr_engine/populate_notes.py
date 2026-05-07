"""Populate noteIDs in extracted CSVs from English reference.

Matches rows by Id to English reference CSVs and copies:
  - noteIDs  ← directly from English reference (language-independent numeric link)

Note_{lang} is NEVER modified — it contains genuine text extracted from
the source PDF and must not be overwritten or translated.

Usage:
    python populate_notes.py \
        --lang fr \
        --input-dir  output/fr/fixed/ \
        --ref-dir    output_en_reference/ \
        --output-dir output/fr/final/
"""

from __future__ import annotations

import argparse
import glob
import os
import re

import pandas as pd


# ── File matching ────────────────────────────────────────────────────────────

def _match_ref_file(lang_file: str, ref_dir: str, lang: str) -> str | None:
    """Find the English reference CSV matching a translated CSV filename."""
    basename = os.path.basename(lang_file)

    # Replace _{lang}_ or _{lang}. with _en_ or _en.
    en_name = re.sub(
        rf"_{lang}([_.])",
        r"_en\1",
        basename,
    )

    # Also handle BUFRCREX_ ↔ BUFR_ prefix swap for Table C
    candidates = [en_name]
    if en_name.startswith("BUFRCREX_TableC_"):
        candidates.append(en_name.replace("BUFRCREX_TableC_", "BUFR_TableC_"))
    elif en_name.startswith("BUFR_TableC_"):
        candidates.append(en_name.replace("BUFR_TableC_", "BUFRCREX_TableC_"))

    for cand in candidates:
        path = os.path.join(ref_dir, cand)
        if os.path.exists(path):
            return path

    return None


# ── Main logic ───────────────────────────────────────────────────────────────

def populate_notes(
    input_dir: str,
    ref_dir: str,
    output_dir: str,
    lang: str,
) -> dict:
    """Populate noteIDs in all CSVs in input_dir from English reference.

    Only copies the noteIDs column (language-independent numeric links).
    Note_{lang} is preserved as-is from the PDF extraction — never
    overwritten or translated.

    Returns summary dict with per-file statistics.
    """
    os.makedirs(output_dir, exist_ok=True)

    summary = {}

    csv_files = sorted(glob.glob(os.path.join(input_dir, "*.csv")))
    if not csv_files:
        print(f"No CSV files found in {input_dir}")
        return summary

    for csv_path in csv_files:
        basename = os.path.basename(csv_path)
        ref_path = _match_ref_file(csv_path, ref_dir, lang)

        if not ref_path:
            # No reference — just copy through
            df = pd.read_csv(csv_path)
            df.to_csv(os.path.join(output_dir, basename), index=False)
            summary[basename] = {"matched": 0, "populated": 0, "total": len(df), "ref": None}
            continue

        df = pd.read_csv(csv_path, dtype=str).fillna("")
        ref = pd.read_csv(ref_path, dtype=str).fillna("")

        if "Id" not in df.columns or "Id" not in ref.columns:
            df.to_csv(os.path.join(output_dir, basename), index=False)
            summary[basename] = {"matched": 0, "populated": 0, "total": len(df), "ref": ref_path}
            continue

        # Build lookup: Id → noteIDs from English reference
        ref_lookup = {}
        for _, row in ref.iterrows():
            rid = row["Id"]
            note_ids = row.get("noteIDs", "")
            if note_ids:
                ref_lookup[rid] = note_ids

        populated = 0
        for idx, row in df.iterrows():
            row_id = row["Id"]
            if row_id in ref_lookup:
                df.at[idx, "noteIDs"] = ref_lookup[row_id]
                populated += 1

        matched = sum(1 for rid in df["Id"] if rid in ref_lookup)

        df.to_csv(os.path.join(output_dir, basename), index=False)
        summary[basename] = {
            "matched": matched,
            "populated": populated,
            "total": len(df),
            "ref": os.path.basename(ref_path),
        }

    return summary


def main():
    p = argparse.ArgumentParser(description="Populate noteIDs from English reference")
    p.add_argument("--lang", required=True, choices=["fr", "es", "ru"])
    p.add_argument("--input-dir", required=True, help="Directory with extracted CSVs")
    p.add_argument("--ref-dir", required=True, help="English reference CSV directory")
    p.add_argument("--output-dir", required=True, help="Output directory")
    args = p.parse_args()

    print(f"Populating noteIDs: {args.lang}")
    print(f"  Input:  {args.input_dir}")
    print(f"  Ref:    {args.ref_dir}")
    print(f"  Output: {args.output_dir}")

    summary = populate_notes(args.input_dir, args.ref_dir, args.output_dir, args.lang)

    total_populated = sum(s["populated"] for s in summary.values())
    total_rows = sum(s["total"] for s in summary.values())
    print(f"\nDone: {len(summary)} files, {total_populated} noteIDs populated across {total_rows} rows")

    for fname, s in sorted(summary.items()):
        if s["populated"] > 0:
            print(f"  {fname}: {s['populated']} noteIDs (matched {s['matched']}/{s['total']} rows)")


if __name__ == "__main__":
    main()
