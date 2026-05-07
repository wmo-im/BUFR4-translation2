"""Compare language CSV files against the English reference.

For each (lang_csv, en_ref_csv) pair matched by filename, this module:
1. Aligns rows on the ``Id`` column.
2. Categorizes every row as MATCH, MISSING_IN_LANG, EXTRA_IN_LANG,
   or ID_MISMATCH (same FXY but different Id string).
3. For CodeFlag files, additionally flags rows where ``CodeFigure``
   implies an ``All-N`` Id but the stored Id has a bare integer.
4. Writes a per-file alignment CSV and a summary report.

Usage
-----
Programmatic:
    from align.align_to_reference import align_directory
    summary = align_directory(lang_dir, ref_dir, output_dir, lang="es")

CLI: see run.py
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import pandas as pd


# ── Helpers ───────────────────────────────────────────────────────────────────

def _read_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, dtype=str, keep_default_na=False)


def _normalize_lang_suffix(col: str, lang: str) -> str:
    """Strip language suffix from a column name to get the base name.

    Examples
    --------
    "ElementName_es" → "ElementName"
    "ElementName_en" → "ElementName"
    """
    return re.sub(rf"_{lang}$|_en$|_fr$|_ru$", "", col)


def _match_filename(lang_name: str, ref_names: list[str], lang: str) -> str | None:
    """Find the English reference filename matching a language filename.

    Handles:
    - Multi-file tables: BUFRCREX_TableB_fr_00.csv → BUFRCREX_TableB_en_00.csv
    - Single-file tables: BUFR_TableA_fr.csv → BUFR_TableA_en.csv
    - Prefix mismatch: BUFRCREX_TableC_fr.csv → BUFR_TableC_en.csv (and vice versa)
    """
    ref_set = set(ref_names)
    # Build all candidate names to try
    candidates = [
        lang_name.replace(f"_{lang}_", "_en_"),   # multi-file (class/category suffix)
        lang_name.replace(f"_{lang}.", "_en."),    # single-file (no trailing _)
    ]
    for c in candidates:
        if c in ref_set:
            return c
        # Try BUFRCREX_ ↔ BUFR_ prefix swap (Table C naming difference)
        if c.startswith("BUFRCREX_"):
            alt = "BUFR_" + c[len("BUFRCREX_"):]
        elif c.startswith("BUFR_"):
            alt = "BUFRCREX_" + c[len("BUFR_"):]
        else:
            alt = c
        if alt in ref_set:
            return alt
    # Last resort: case-insensitive stem match
    lang_stem = Path(lang_name).stem.upper()
    lang_stem_norm = lang_stem.replace(f"_{lang.upper()}_", "_EN_").replace(f"_{lang.upper()}", "_EN")
    for rn in ref_names:
        if Path(rn).stem.upper() == lang_stem_norm:
            return rn
    return None


# ── Core alignment ────────────────────────────────────────────────────────────

_ALL_N_PATTERN = re.compile(r"^(?i:todo[s]?|los|all)\s*[-\u2013]?\s*\d+$")


def align_file(
    lang_path: Path,
    ref_path: Path,
    lang: str,
) -> pd.DataFrame:
    """Align one language CSV against its English reference.

    Parameters
    ----------
    lang_path : Path
        Language CSV file (after ID validation/fixing).
    ref_path : Path
        Corresponding English reference CSV.
    lang : str
        Two-letter language code (used for column normalization).

    Returns
    -------
    pd.DataFrame with columns:
        Id, Status, lang_row, ref_row  — one row per unique Id
        where Status ∈ {MATCH, MISSING_IN_LANG, EXTRA_IN_LANG, ID_MISMATCH,
                        ALL_N_MISMATCH}
    """
    lang_df = _read_csv(lang_path)
    ref_df = _read_csv(ref_path)

    is_codeflag = "codeflag" in lang_path.stem.lower() or "code_flag" in lang_path.stem.lower()

    lang_ids = set(lang_df["Id"]) if "Id" in lang_df.columns else set()
    ref_ids = set(ref_df["Id"]) if "Id" in ref_df.columns else set()

    rows = []

    # MISSING_IN_LANG: present in reference, absent in language file
    for id_val in sorted(ref_ids - lang_ids):
        rows.append({
            "Id": id_val,
            "Status": "MISSING_IN_LANG",
            "lang_row": None,
            "ref_row": _find_row(ref_df, id_val),
        })

    # EXTRA_IN_LANG: present in language file, absent in reference
    for id_val in sorted(lang_ids - ref_ids):
        # For CodeFlag: check if it's an All-N Id without dash (bare integer where All-N expected)
        status = "EXTRA_IN_LANG"
        if is_codeflag:
            cf_seg = id_val.split("/")[-1] if "/" in id_val else ""
            # If CodeFigure col in lang_df has All-N text but Id doesn't use All-N
            lang_row = _find_row(lang_df, id_val)
            if lang_row is not None:
                cf_val = str(lang_row.get("CodeFigure", "")).strip()
                if _ALL_N_PATTERN.match(cf_val) and not cf_seg.startswith("All"):
                    status = "ALL_N_MISMATCH"
        rows.append({
            "Id": id_val,
            "Status": status,
            "lang_row": _find_row(lang_df, id_val),
            "ref_row": None,
        })

    # MATCH or ID_MISMATCH for shared Ids
    for id_val in sorted(lang_ids & ref_ids):
        rows.append({
            "Id": id_val,
            "Status": "MATCH",
            "lang_row": _find_row(lang_df, id_val),
            "ref_row": _find_row(ref_df, id_val),
        })

    return pd.DataFrame(rows)


def _find_row(df: pd.DataFrame, id_val: str) -> dict | None:
    """Return the first row with Id == id_val as a dict, or None."""
    if "Id" not in df.columns:
        return None
    mask = df["Id"] == id_val
    if not mask.any():
        return None
    return df[mask].iloc[0].to_dict()


# ── Directory-level alignment ─────────────────────────────────────────────────

def align_directory(
    lang_dir: Path,
    ref_dir: Path,
    output_dir: Path,
    lang: str,
) -> list[dict]:
    """Align all language CSVs in *lang_dir* against English reference in *ref_dir*.

    Writes per-file alignment CSVs to *output_dir* and returns a summary list.

    Parameters
    ----------
    lang_dir : Path
        Directory containing fixed language CSVs.
    ref_dir : Path
        Directory containing English reference CSVs.
    output_dir : Path
        Directory to write alignment result CSVs.
    lang : str
        Two-letter language code.

    Returns
    -------
    list of summary dicts:
        {file, ref_file, lang_rows, ref_rows, match, missing, extra, all_n_mismatch}
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    lang_files = sorted(lang_dir.glob("*.csv"))
    ref_names = [f.name for f in ref_dir.glob("*.csv")]
    summary = []

    for lang_path in lang_files:
        ref_name = _match_filename(lang_path.name, ref_names, lang)
        if ref_name is None:
            print(f"  [SKIP] No reference match for {lang_path.name}")
            summary.append({
                "file": lang_path.name, "ref_file": None,
                "lang_rows": None, "ref_rows": None,
                "match": 0, "missing": 0, "extra": 0, "all_n_mismatch": 0,
                "note": "No reference file found",
            })
            continue

        ref_path = ref_dir / ref_name
        try:
            result_df = align_file(lang_path, ref_path, lang)
        except Exception as exc:
            print(f"  [ERROR] {lang_path.name}: {exc}")
            summary.append({
                "file": lang_path.name, "ref_file": ref_name,
                "lang_rows": None, "ref_rows": None,
                "match": 0, "missing": 0, "extra": 0, "all_n_mismatch": 0,
                "note": str(exc),
            })
            continue

        # Count statuses
        status_counts = result_df["Status"].value_counts().to_dict()
        n_match = status_counts.get("MATCH", 0)
        n_miss = status_counts.get("MISSING_IN_LANG", 0)
        n_extra = status_counts.get("EXTRA_IN_LANG", 0)
        n_all_n = status_counts.get("ALL_N_MISMATCH", 0)

        # Write alignment CSV
        out_name = f"alignment_{lang_path.stem}.csv"
        # Flatten the ref_row/lang_row dicts to individual columns for readability
        flat_rows = []
        for _, row in result_df.iterrows():
            flat = {"Id": row["Id"], "Status": row["Status"]}
            # Add key fields from the ref row if available
            ref_row = row.get("ref_row")
            if isinstance(ref_row, dict):
                flat["ref_Id"] = ref_row.get("Id", "")
            flat_rows.append(flat)

        pd.DataFrame(flat_rows).to_csv(output_dir / out_name, index=False)

        lang_rows_count = len(_read_csv(lang_path))
        ref_rows_count = len(_read_csv(ref_path))

        coverage = (
            f"{100 * n_match / ref_rows_count:.1f}%"
            if ref_rows_count > 0 else "n/a"
        )
        status = "OK" if n_miss == 0 and n_extra == 0 and n_all_n == 0 else "DIFF"
        print(
            f"  {lang_path.name}: {n_match}/{ref_rows_count} matched ({coverage}), "
            f"{n_miss} missing, {n_extra} extra, {n_all_n} All-N mismatch [{status}]"
        )

        summary.append({
            "file": lang_path.name,
            "ref_file": ref_name,
            "lang_rows": lang_rows_count,
            "ref_rows": ref_rows_count,
            "match": n_match,
            "missing": n_miss,
            "extra": n_extra,
            "all_n_mismatch": n_all_n,
            "coverage_pct": coverage,
            "note": status,
        })

    return summary


def write_summary_report(summary: list[dict], report_path: Path) -> None:
    """Write a markdown summary report from the alignment summary list."""
    lines = [
        "# BUFR Alignment Report",
        "",
        "| File | Ref File | Lang Rows | Ref Rows | Match | Missing | Extra | All-N | Coverage |",
        "|------|----------|-----------|----------|-------|---------|-------|-------|----------|",
    ]
    totals = {"lang_rows": 0, "ref_rows": 0, "match": 0, "missing": 0,
               "extra": 0, "all_n_mismatch": 0}

    for row in summary:
        lines.append(
            f"| {row['file']} | {row.get('ref_file', '-')} | "
            f"{row.get('lang_rows', '-')} | {row.get('ref_rows', '-')} | "
            f"{row.get('match', 0)} | {row.get('missing', 0)} | "
            f"{row.get('extra', 0)} | {row.get('all_n_mismatch', 0)} | "
            f"{row.get('coverage_pct', '-')} |"
        )
        for k in totals:
            v = row.get(k, 0)
            if isinstance(v, int):
                totals[k] += v

    lines.extend([
        "",
        f"**Totals**: {totals['lang_rows']} lang rows, {totals['ref_rows']} ref rows | "
        f"{totals['match']} matched | {totals['missing']} missing | "
        f"{totals['extra']} extra | {totals['all_n_mismatch']} All-N mismatches",
        "",
    ])

    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"Report written: {report_path}")
