"""ID validation, reconstruction, and normalization for BUFR table CSVs.

All BUFR table rows carry an ``Id`` column with a strict format:
  - Table A:       bufr4/a/{CodeFigure}
  - Table B:       bufr4/{F}/{XX}/{YYY}
  - Table C:       bufr4/{FXY}               (symbolic, e.g. 201YYY)
  - Table D:       bufr4/3/{XX}/{YYY}/{occ}/{FXY2:06d}
  - CodeFlag:      bufr4/{F}/{XX}/{YYY}/{cf}

This module is language-independent: it derives correct Ids from the
numeric/symbolic columns already present in the CSV, so no English
reference is needed.

Public API
----------
  is_valid_id(id_str)            → bool
  detect_table_type(stem)        → str
  reconstruct_id(row, type, occ) → str | None
  fix_file(path, output_dir)     → (fixed_df, changes_list)
"""

from __future__ import annotations

import re
import pathlib
import pandas as pd


# ── Valid ID segment patterns ─────────────────────────────────────────────────
# After splitting "bufr4/<rest>" on '/', each segment must match one of:
#   \d+            pure digits        e.g. "002", "000001"
#   a              Table A prefix
#   \d+[A-Z]+      digits+uppercase   e.g. "201YYY", "C01YYY"
#   [A-Z]+         pure uppercase     e.g. "YYY", "XX"
#   \d+[-–]\d+     numeric range      e.g. "8-30", "15–19"   (reserved entries)
#   All[-–]\d+     WMO missing-value  e.g. "All-16"
#
# Anything with lowercase letters NOT fitting the above is invalid.
# Specifically, this rejects Spanish words like "Todo-16" or "continúa".
_VALID_SEG = re.compile(
    r"^\d+$"              # pure digits
    r"|^a$"               # Table A single-letter prefix
    r"|^\d+[A-Z]+$"       # digits + uppercase  (201YYY)
    r"|^[A-Z]+$"          # pure uppercase  (YYY, XX)
    r"|^\d+[-\u2013]\d+$" # numeric range: 8-30 or 15–19 (en-dash U+2013)
    r"|^All[-\u2013]\d+$" # WMO All-N convention (missing values)
)

# Force-flag pattern for CodeFlag rows: CodeFigure contains Spanish missing-value
# labels that produce syntactically valid but semantically wrong Ids.
# "Todo N" / "Todos N" / "Los N" → should be "All-N" in the Id.
_ALL_N_CF = re.compile(
    r"^(?:tou(?:s|tes?)|todo[s]?|los|all)\s+\d+$"
    r"|^\d+\s+bits?\s+mis\s+à\s+1$",
    re.IGNORECASE,
)


def is_valid_id(id_str: str) -> bool:
    """Return True iff *id_str* is a well-formed bufr4/... identifier.

    Accepts range-type code-figure segments (``8-30``, ``All-16``) that are
    standard WMO BUFR table entries. Rejects segments with unexpected lowercase
    text (e.g. ``Todo-16``, ``continúa``).
    """
    if not isinstance(id_str, str):
        return False
    if not id_str.startswith("bufr4/"):
        return False
    rest = id_str[len("bufr4/"):]
    if not rest:
        return False
    for seg in rest.split("/"):
        if not _VALID_SEG.match(seg):
            return False
    return True


# ── Table-type detection from filename stem ───────────────────────────────────

def detect_table_type(stem: str) -> str:
    """Infer table type from filename stem.

    Returns one of: ``table_a``, ``table_b``, ``table_c``, ``table_d``,
    ``codeflag``, or ``unknown``.

    Examples
    --------
    >>> detect_table_type("BUFRCREX_TableB_es_00")
    'table_b'
    >>> detect_table_type("BUFRCREX_CodeFlag_es_02")
    'codeflag'
    """
    s = stem.upper()
    if "TABLEA" in s or "TABLE_A" in s:
        return "table_a"
    if "TABLEC" in s or "TABLE_C" in s:
        return "table_c"
    if "TABLED" in s or "TABLE_D" in s:
        return "table_d"
    if "CODEFLAG" in s or "CODE_FLAG" in s:
        return "codeflag"
    if "TABLEB" in s or "TABLE_B" in s:
        return "table_b"
    return "unknown"


# ── ID reconstruction ─────────────────────────────────────────────────────────

def _fxy_to_parts(fxy_int: int) -> tuple[str, str, str]:
    """Decompose an integer FXY code into (F, XX, YYY) zero-padded strings."""
    s = f"{int(fxy_int):06d}"
    return s[0], s[1:3], s[3:6]


def reconstruct_id(row: pd.Series, table_type: str, occ: int = 1) -> str | None:
    """Build the correct Id string from FXY/CodeFigure columns.

    Parameters
    ----------
    row : pd.Series
        A single CSV row (from ``df.iterrows()``).
    table_type : str
        One of the values returned by ``detect_table_type()``.
    occ : int
        1-based occurrence index within the FXY1 group (Table D only).
        Pre-compute with ``df.groupby("FXY1").cumcount() + 1``.

    Returns
    -------
    str | None
        Reconstructed Id, or None if required columns are missing/non-numeric.

    Notes
    -----
    For CodeFlag, ``CodeFigure`` is normalized:
    - ``"Todo N"`` / ``"Todos N"`` / ``"Los N"`` → ``"All-N"``
    - ``"All N"`` (space) → ``"All-N"`` (hyphen)
    - Integer values → plain string (``"6"``)
    - Range values kept as-is (``"8-30"``, ``"All-16"``)
    """
    try:
        if table_type == "table_a":
            cf = int(row["CodeFigure"])
            return f"bufr4/a/{cf}"

        elif table_type == "table_b":
            fxy = int(row["FXY"])
            F, XX, YYY = _fxy_to_parts(fxy)
            return f"bufr4/{F}/{XX}/{YYY}"

        elif table_type == "table_c":
            fxy = str(row["FXY"]).strip()
            if not fxy:
                return None
            return f"bufr4/{fxy}"

        elif table_type == "table_d":
            fxy1 = int(row["FXY1"])
            fxy2 = int(row["FXY2"])
            _, XX, YYY = _fxy_to_parts(fxy1)  # F is always 3 for Table D
            fxy2_str = f"{fxy2:06d}"
            return f"bufr4/3/{XX}/{YYY}/{occ}/{fxy2_str}"

        elif table_type == "codeflag":
            fxy = int(row["FXY"])
            F, XX, YYY = _fxy_to_parts(fxy)
            cf_raw = str(row["CodeFigure"]).strip()
            # Normalize Spanish/French missing-value labels → WMO "All-N" convention
            # French: "4 bits mis à 1" → "All-4"
            cf = re.sub(r"^(\d+)\s+bits?\s+mis\s+à\s+1$", r"All-\1", cf_raw, flags=re.IGNORECASE)
            cf = re.sub(r"^(?:tou(?:s|tes?)|todo[s]?|los|все)\s+", "All-", cf, flags=re.IGNORECASE)
            cf = re.sub(r"^(?:(?i:todo[s]?)|все)[-\u2013]", "All-", cf)
            cf = re.sub(r"^All\s+", "All-", cf)  # "All N" (space) → "All-N"
            try:
                cf = str(int(cf))  # integer code figures → plain string
            except ValueError:
                pass  # keep range strings: "8-30", "All-16"
            return f"bufr4/{F}/{XX}/{YYY}/{cf}"

    except (ValueError, TypeError, KeyError):
        return None

    return None


# ── Per-file processing ───────────────────────────────────────────────────────

def fix_file(
    path: pathlib.Path,
    output_dir: pathlib.Path | None = None,
) -> tuple[pd.DataFrame, list[dict]]:
    """Validate and fix the Id column in one BUFR CSV file.

    Parameters
    ----------
    path : pathlib.Path
        Input CSV path.
    output_dir : pathlib.Path | None
        If given, writes the fixed CSV to this directory under the same filename.
        If None, returns the fixed DataFrame without writing.

    Returns
    -------
    (fixed_df, changes_list)
        ``fixed_df`` — DataFrame with corrected Ids.
        ``changes_list`` — list of dicts with keys:
            file, row_index, old_id, new_id, method

    Notes
    -----
    Always reads with ``dtype=str`` to preserve leading zeros in FXY columns.
    """
    df = pd.read_csv(path, dtype=str)
    stem = path.stem
    table_type = detect_table_type(stem)
    changes: list[dict] = []

    # Pre-compute per-(FXY1, FXY2) occurrence counters for Table D
    occ_series: pd.Series | None = None
    if table_type == "table_d" and "FXY1" in df.columns and "FXY2" in df.columns:
        occ_series = df.groupby(["FXY1", "FXY2"]).cumcount() + 1

    for idx, row in df.iterrows():
        old_id = row.get("Id", None)
        if pd.isna(old_id):
            old_id = ""

        # Force-flag: "Los N" / "Todo N" in CodeFigure produces a syntactically
        # valid but semantically wrong Id (plain digit instead of "All-N").
        force = (
            table_type == "codeflag"
            and "CodeFigure" in row.index
            and _ALL_N_CF.match(str(row["CodeFigure"]).strip())
        )

        if is_valid_id(str(old_id)) and not force:
            continue

        occ = int(occ_series.loc[idx]) if occ_series is not None else 1
        new_id = reconstruct_id(row, table_type, occ=occ)
        method = "reconstructed"

        if new_id is None:
            method = "unresolved"
            changes.append({
                "file": path.name, "row_index": idx,
                "old_id": old_id, "new_id": "", "method": method,
            })
            continue

        df.at[idx, "Id"] = new_id
        changes.append({
            "file": path.name, "row_index": idx,
            "old_id": old_id, "new_id": new_id, "method": method,
        })

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        df.to_csv(output_dir / path.name, index=False)

    return df, changes


def fix_directory(
    input_dir: pathlib.Path,
    output_dir: pathlib.Path,
) -> tuple[list[dict], int]:
    """Fix all CSVs in *input_dir* and write results to *output_dir*.

    Returns
    -------
    (all_changes, file_count)
    """
    csv_files = sorted(input_dir.glob("*.csv"))
    all_changes: list[dict] = []

    for path in csv_files:
        _, changes = fix_file(path, output_dir)
        all_changes.extend(changes)
        if changes:
            n_ok = sum(1 for c in changes if c["method"] != "unresolved")
            n_bad = sum(1 for c in changes if c["method"] == "unresolved")
            print(f"  {path.name}: {n_ok} fixed, {n_bad} unresolved")
        else:
            print(f"  {path.name}: OK")

    return all_changes, len(csv_files)
