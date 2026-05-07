#!/usr/bin/env python3
"""
Align extracted CCT CSV files against English reference.

For each table x language:
1. Read English reference (correct structure, IDs, code figures)
2. Read freshly extracted CSV (has translated text, but header junk and shifted rows)
3. Build code-figure lookup from extracted data, parsing codes from wherever they appear
4. Reconstruct aligned output using English skeleton + translated text
"""

import csv
import io
import re
import sys
from pathlib import Path
from collections import defaultdict, OrderedDict

EN_DIR = Path("/Users/omard/Documents/projects/CCT-translation/english")
BASE_DIR = Path("/Users/omard/Documents/projects/WMO_work_claude/table_extractor_vision/wmo_pipeline/data")
LANGUAGES = ["ru", "fr", "es"]
TABLES = ["C00", "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08",
          "C11", "C12", "C13", "C14"]


# ── Helpers ──────────────────────────────────────────────────────────────────

def read_csv(path):
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    # Fix C11 English header (malformed - missing newline after translation_source)
    if "translation_sourceC11" in text:
        text = text.replace("translation_sourceC11", "translation_source\nC11")
    reader = csv.DictReader(io.StringIO(text))
    return reader.fieldnames, list(reader)


def write_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def ndash(s):
    """Normalize dashes and strip."""
    if not s:
        return s
    return (s.replace("\u2010", "-").replace("\u2011", "-").replace("\u2012", "-")
             .replace("\u2013", "-").replace("\u2014", "-").replace("\u00ad", "").strip())


def fix_formula(f):
    """Fix subscript spacing in chemical formulas."""
    if not f:
        return f
    fixes = {
        "O 3": "O3", "H O 2": "H2O", "H2 O": "H2O", "CH 4": "CH4",
        "CO 2": "CO2", "NO 2": "NO2", "N O 2": "N2O", "N2 O": "N2O",
        "N 2O": "N2O", "SO 2": "SO2", "NH 3": "NH3",
        "CFC 11": "CFC-11", "CFC 12": "CFC-12", "CFC 113": "CFC-113",
        "CFC 114": "CFC-114", "CFC 22": "CFC-22",
        "CCl 4": "CCl4", "CF 4": "CF4",
        "CCl 3F": "CCl3F", "CCl 2F 2": "CCl2F2",
        "CHCl F 2": "CHClF2",
        "CCl 2FCClF 2": "CCl2FCClF2", "CClF 2CClF 2": "CClF2CClF2",
    }
    s = f.strip()
    return fixes.get(s, s)


def strip_leading_number(text):
    """Strip leading integer from text like '1 Melbourne' -> 'Melbourne'. Returns (num_str, rest)."""
    if not text:
        return None, text
    m = re.match(r'^(\d+)\s+(.+)$', text.strip())
    if m:
        return m.group(1), m.group(2)
    return None, text.strip()


def is_junk(text):
    """Detect header-junk text that leaked from column headers."""
    if not text:
        return False
    t = text.strip().lower()
    patterns = [
        "кодовая", "цифра", "обычное", "сокращение", "определение",
        "для f", "pour f", "para f", "pour x", "para x",
        "pour s", "para s", "pour i", "para i",
        "table de code", "tabla de", "común",
        "édition", "edición", "octet", "section",
        "nécessaire", "necesaria",
        "du code", "clave", "alfanuméricas", "alphanu",
        "crex édition", "crex edición", "bufr édition", "grib édition",
        "crex edición", "bufr edición", "grib edición",
        "caractères", "(table de", "(tabla de",
        "catégories", "categorías", "ous -",
        "sous-catég", "subcateg",
        "c entres", "centros de origen",
        "r égion", "р егион",
        "traditionnel", "tradicional",
        "x x", "¡", "для r r",
        "chiffre de code", "cifra de",
    ]
    for p in patterns:
        if p in t:
            return True
    return False


def en_col_to_lang(col, lang):
    """OriginatingGeneratingCentres_en -> OriginatingGeneratingCentres_ru"""
    if col.endswith("_en"):
        return col[:-3] + f"_{lang}"
    return col


def make_out_headers(en_headers, lang):
    return [en_col_to_lang(h, lang) for h in en_headers]


# ── Per-table alignment ─────────────────────────────────────────────────────

def align_c00(lang):
    """C00: no translatable columns, just copy English."""
    en_h, en_rows = read_csv(EN_DIR / "C00_en.csv")
    out_path = BASE_DIR / lang / "cct_output" / f"C00_{lang}.csv"
    write_csv(out_path, en_h, en_rows)
    return len(en_rows), len(en_rows), 0


def _generic_align(table, lang, code_getter_en, code_getter_ext, text_cols_en,
                   extra_copy_cols=None, skip_header_rows=0, formula_col=None):
    """
    Generic alignment engine.

    code_getter_en(en_row) -> str or None (the code figure to match on)
    code_getter_ext(ext_row, lang) -> (code_str, dict_of_lang_text_cols)
    text_cols_en: list of English column names ending in _en that contain translatable text
    extra_copy_cols: non-text cols to copy from English (e.g. 'conventional', 'NoteID')
    """
    en_h, en_rows = read_csv(EN_DIR / f"{table}_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"{table}_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_cols_lang = [en_col_to_lang(c, lang) for c in text_cols_en]
    out_headers = make_out_headers(en_h, lang)

    # Build lookup from extracted data
    lookup = OrderedDict()  # code -> {lang_col: text}
    group_texts = []  # for group header rows

    for i, row in enumerate(ext_rows):
        if i < skip_header_rows:
            continue

        code, texts = code_getter_ext(row, lang)
        if code is None:
            continue

        # Skip junk
        all_text = " ".join(str(v) for v in texts.values() if v)
        if is_junk(all_text):
            continue

        # Group header detection (text like "01-09: WMCs")
        if not code and texts:
            first_text = next((v for v in texts.values() if v), "")
            if re.match(r'^\d+-\d+:', first_text.strip()):
                group_texts.append(ndash(first_text.strip()))
            continue

        if code:
            lookup[code] = {k: ndash(str(v)) for k, v in texts.items()}

    # Reconstruct
    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = code_getter_en(en_row)
        en_id = en_row.get("ID", "")

        out = {}
        # Copy all non-_en columns from English
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")

        # Always set Status from English
        out["Status"] = en_row.get("Status", "Operational")

        ext = lookup.get(en_code) if en_code else None

        if ext:
            for tc_en, tc_lang in zip(text_cols_en, text_cols_lang):
                out[tc_lang] = ext.get(tc_lang, "")
            out["translation_source"] = "PDF"
            matched += 1
        elif "//" in str(en_id):
            # Group header - try matching range text
            en_text = str(en_row.get(text_cols_en[0], "")).strip()
            found_group = ""
            range_m = re.match(r'^(\d+-\d+)', en_text)
            if range_m:
                prefix = range_m.group(1)
                for gt in group_texts:
                    if gt.startswith(prefix):
                        found_group = gt
                        break
            for tc_lang in text_cols_lang:
                out[tc_lang] = found_group if tc_lang == text_cols_lang[0] else ""
            out["translation_source"] = "PDF" if found_group else "NA"
            if found_group:
                matched += 1
            else:
                unmatched += 1
        else:
            for tc_lang in text_cols_lang:
                out[tc_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        # Copy extra cols from English
        if extra_copy_cols:
            for ec in extra_copy_cols:
                if ec in en_row:
                    out[ec] = en_row[ec]

        # Fix formula if needed
        if formula_col and formula_col in en_row:
            out[formula_col] = en_row[formula_col]

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c01(lang):
    """C01: match by F1F2 (00-99) or F3F3F3 (100+)."""
    en_h, en_rows = read_csv(EN_DIR / "C01_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C01_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_en = "OriginatingGeneratingCentres_en"
    text_col_lang = f"OriginatingGeneratingCentres_{lang}"
    out_headers = make_out_headers(en_h, lang)

    # Build lookup: try to find the code figure in each extracted row
    # Code can be in CodeFigureForF1F2, CodeFigureForF3F3F3, or Octet5
    lookup = {}  # code_key -> translated_text
    group_texts = []

    for row in ext_rows:
        f1f2 = str(row.get("CodeFigureForF1F2", "")).strip()
        f3 = str(row.get("CodeFigureForF3F3F3", "")).strip()
        octet = str(row.get("Octet5GRIB1_Octet6BUFR3", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()

        # Skip junk rows
        if is_junk(f"{f1f2} {f3} {octet}"):
            continue

        # Group header detection
        if text and re.match(r'^\d+-\d+:', text):
            group_texts.append(ndash(text))
            continue

        # Try to determine the code figure
        code_key = None

        # Best: use F3F3F3 (3-digit, unique)
        if re.match(r'^\d{3}$', f3):
            code_key = f3
        # Fallback: F1F2 (2-digit)
        elif re.match(r'^\d{2}$', f1f2):
            code_key = f1f2.lstrip("0") or "0"  # normalize "00" -> "0" for matching
            # Actually keep as-is for lookup, we'll match both ways
            code_key = f1f2

        if code_key and text:
            # Store by F3F3F3 for reliable matching
            if re.match(r'^\d{3}$', f3):
                lookup[f3] = ndash(text)
            elif re.match(r'^\d{2}$', f1f2):
                lookup[f1f2] = ndash(text)

    # Reconstruct
    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_f1f2 = str(en_row.get("CodeFigureForF1F2", "")).strip()
        en_f3 = str(en_row.get("CodeFigureForF3F3F3", "")).strip()
        en_id = en_row.get("ID", "")
        en_text = str(en_row.get(text_col_en, "")).strip()

        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        # Try match by F3F3F3 first, then F1F2
        text = lookup.get(en_f3)
        if text is None and re.match(r'^\d{2}$', en_f1f2):
            text = lookup.get(en_f1f2)

        if text is not None:
            out[text_col_lang] = text
            out["translation_source"] = "PDF"
            matched += 1
        elif "//" in str(en_id):
            found = ""
            range_m = re.match(r'^(\d+-\d+)', en_text)
            if range_m:
                for gt in group_texts:
                    if gt.startswith(range_m.group(1)):
                        found = gt
                        break
            # Also check for non-range group headers like "Additional Centres"
            out[text_col_lang] = found
            out["translation_source"] = "PDF" if found else "NA"
            if found:
                matched += 1
            else:
                unmatched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c02(lang):
    """
    C02: Code figures are in CodeFigureForBUFR column.
    Russian: "XX YY" (rara + BUFR merged), French/Spanish: just rara "XX" with BUFR in text.
    """
    en_h, en_rows = read_csv(EN_DIR / "C02_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C02_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    date_col_lang = f"DateOfAssignment_{lang}"
    text_col_lang = f"RadiosondeSoundingSystemUsed_{lang}"
    out_headers = make_out_headers(en_h, lang)

    # Build lookup by rara code (2-digit)
    lookup = {}  # rara_code -> (date_text, sounding_text)

    for row in ext_rows:
        bufr_val = str(row.get("CodeFigureForBUFR", "")).strip()
        date_val = str(row.get(date_col_lang, "")).strip()
        text_val = str(row.get(text_col_lang, "")).strip()

        # Skip header junk
        if is_junk(bufr_val):
            continue

        # Parse code figure
        rara_code = None

        # Russian: "XX YY" in BUFR column (rara + BUFR merged)
        m = re.match(r'^(\d{2})\s+(\d+)$', bufr_val)
        if m:
            rara_code = m.group(1)
        # French/Spanish: just rara "XX" in BUFR column, BUFR code prepended to text
        elif re.match(r'^\d{2}$', bufr_val):
            rara_code = bufr_val
            # Strip leading BUFR code from text
            num, rest = strip_leading_number(text_val)
            if num is not None:
                text_val = rest

        if rara_code and not is_junk(text_val):
            lookup[rara_code] = (ndash(date_val), ndash(text_val))

    # Reconstruct
    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_rara = str(en_row.get("CodeFigureForrara", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        data = lookup.get(en_rara)
        if data:
            out[date_col_lang] = data[0]
            out[text_col_lang] = data[1]
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[date_col_lang] = ""
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c03(lang):
    """C03: CodeFigureForIXIIXIX is the key (3-digit)."""
    en_h, en_rows = read_csv(EN_DIR / "C03_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C03_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_lang = f"InstrumentMakeAndType_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}  # IX code -> text
    for row in ext_rows:
        code = str(row.get("CodeFigureForIXIIXIX", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()
        bufr = str(row.get("CodeFigureForBUFR", "")).strip()

        if is_junk(f"{code} {bufr}"):
            continue

        # Text often has "BUFR_code InstrumentName" prepended
        num, rest = strip_leading_number(text)
        if num is not None:
            text = rest

        if re.match(r'^\d{3}$', code):
            lookup[code] = ndash(text)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = str(en_row.get("CodeFigureForIXIIXIX", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        text = lookup.get(en_code)
        if text is not None:
            out[text_col_lang] = text
            # Always use English coefficients
            out["EquationCoefficients_a"] = en_row.get("EquationCoefficients_a", "")
            out["EquationCoefficients_b"] = en_row.get("EquationCoefficients_b", "")
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c04(lang):
    """C04: CodeFigureForXRXR (2-digit)."""
    en_h, en_rows = read_csv(EN_DIR / "C04_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C04_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_lang = f"Meaning_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    for row in ext_rows:
        code = str(row.get("CodeFigureForXRXR", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()

        if is_junk(code):
            continue
        num, rest = strip_leading_number(text)
        if num is not None:
            text = rest
        if re.match(r'^\d{2}$', code):
            lookup[code] = ndash(text)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = str(en_row.get("CodeFigureForXRXR", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        text = lookup.get(en_code)
        if text is not None:
            out[text_col_lang] = text
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c05(lang):
    """C05: CodeFigureForI6I6I6 (3-digit) or CodeFigureForBUFR."""
    en_h, en_rows = read_csv(EN_DIR / "C05_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C05_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_en = "SatelliteName_en"
    text_col_lang = f"SatelliteName_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    group_texts = []

    for row in ext_rows:
        i6 = str(row.get("CodeFigureForI6I6I6", "")).strip()
        bufr = str(row.get("CodeFigureForBUFR", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()

        if is_junk(f"{i6} {bufr}"):
            continue

        # Group headers
        if text and re.match(r'^\d+-\d+:', text):
            group_texts.append(ndash(text))
            continue

        # Text often has GRIB2 code prepended
        num, rest = strip_leading_number(text)
        if num is not None:
            text = rest

        if re.match(r'^\d{3}$', i6):
            lookup[i6] = ndash(text)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = str(en_row.get("CodeFigureForI6I6I6", "")).strip()
        en_id = en_row.get("ID", "")
        en_text = str(en_row.get(text_col_en, "")).strip()

        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        text = lookup.get(en_code)
        if text is not None:
            out[text_col_lang] = text
            out["translation_source"] = "PDF"
            matched += 1
        elif "//" in str(en_id):
            found = ""
            range_m = re.match(r'^(\d+-\d+)', en_text)
            if range_m:
                for gt in group_texts:
                    if gt.startswith(range_m.group(1)):
                        found = gt
                        break
            out[text_col_lang] = found
            out["translation_source"] = "PDF" if found else "NA"
            if found:
                matched += 1
            else:
                unmatched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c06(lang):
    """C06: CodeFigure (3-digit). Text cols: UnitType, Meaning."""
    en_h, en_rows = read_csv(EN_DIR / "C06_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C06_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    ut_lang = f"UnitType_{lang}"
    m_lang = f"Meaning_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    for row in ext_rows:
        code = str(row.get("CodeFigure", "")).strip()
        ut = str(row.get(ut_lang, "")).strip()
        mg = str(row.get(m_lang, "")).strip()

        if is_junk(code):
            continue
        if re.match(r'^\d{3}$', code):
            lookup[code] = (ndash(ut), ndash(mg))

    out_rows = []
    matched = unmatched = 0
    tech_cols = ["conventional", "IA5-ASCII", "ITA2", "SIDefinition", "Note", "NoteID"]

    for en_row in en_rows:
        en_code = str(en_row.get("CodeFigure", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")
        for tc in tech_cols:
            if tc in en_row:
                out[tc] = en_row[tc]

        data = lookup.get(en_code)
        if data:
            out[ut_lang] = data[0]
            out[m_lang] = data[1]
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[ut_lang] = ""
            out[m_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c07(lang):
    """C07: CodeFigureForsasa (2-digit)."""
    en_h, en_rows = read_csv(EN_DIR / "C07_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C07_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_lang = f"TrackingTechniquesStatusOfSystemUsed_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    for row in ext_rows:
        code = str(row.get("CodeFigureForsasa", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()

        if is_junk(code):
            continue
        num, rest = strip_leading_number(text)
        if num is not None:
            text = rest
        if re.match(r'^\d{2}$', code):
            lookup[code] = ndash(text)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = str(en_row.get("CodeFigureForsasa", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        text = lookup.get(en_code)
        if text is not None:
            out[text_col_lang] = text
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c08(lang):
    """C08: Code (numeric). Multiple text cols."""
    en_h, en_rows = read_csv(EN_DIR / "C08_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C08_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_cols = [f"Agency_{lang}", f"Type_{lang}",
                 f"InstrumentShortName_{lang}", f"InstrumentLongName_{lang}"]
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    for row in ext_rows:
        code = str(row.get("Code", "")).strip()
        if not re.match(r'^\d+$', code):
            continue
        if is_junk(" ".join(str(row.get(c, "")) for c in text_cols)):
            continue
        lookup[code] = {c: ndash(str(row.get(c, ""))) for c in text_cols}

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = str(en_row.get("Code", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        data = lookup.get(en_code)
        if data:
            for c in text_cols:
                out[c] = data.get(c, "")
            out["translation_source"] = "PDF"
            matched += 1
        else:
            for c in text_cols:
                out[c] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c11(lang):
    """C11: CREX2 (5-digit). Text: OriginatingGeneratingCentre."""
    en_h, en_rows = read_csv(EN_DIR / "C11_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C11_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_en = "OriginatingGeneratingCentre_en"
    text_col_lang = f"OriginatingGeneratingCentre_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    group_texts = []

    for row in ext_rows:
        crex = str(row.get("CREX2", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()

        if is_junk(crex):
            continue

        # Handle merged text like "Secrétariat de l'OMM 00001-00009: CMM"
        if text and re.search(r'\d{5}-\d{5}:', text):
            m = re.match(r'^(.+?)\s+(\d{5}-\d{5}:.+)$', text)
            if m:
                text = m.group(1)
                group_texts.append(ndash(m.group(2)))

        if text and re.match(r'^\d+-\d+:', text):
            group_texts.append(ndash(text))
            continue

        if re.match(r'^\d{5}$', crex):
            lookup[crex] = ndash(text)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_crex = str(en_row.get("CREX2", "")).strip()
        en_id = en_row.get("ID", "")
        en_text = str(en_row.get(text_col_en, "")).strip()

        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        text = lookup.get(en_crex)
        if text is not None:
            out[text_col_lang] = text
            out["translation_source"] = "PDF"
            matched += 1
        elif "//" in str(en_id):
            found = ""
            range_m = re.match(r'^(\d+-\d+)', en_text)
            if range_m:
                for gt in group_texts:
                    # Normalize for matching: strip leading zeros
                    gt_range = re.match(r'^(\d+)', gt)
                    en_range = range_m.group(1)
                    if gt_range:
                        gt_start = gt_range.group(1).lstrip("0") or "0"
                        en_start = en_range.split("-")[0].lstrip("0") or "0"
                        if gt_start == en_start:
                            found = gt
                            break
            out[text_col_lang] = found
            out["translation_source"] = "PDF" if found else "NA"
            if found:
                matched += 1
            else:
                unmatched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c12(lang):
    """C12: compound key (OriginatingCentres + SubCentres codes)."""
    en_h, en_rows = read_csv(EN_DIR / "C12_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C12_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    oc_code = "CodeFigure_OriginatingCentres"
    sc_code = "CodeFigure_SubCentres"
    oc_name_lang = f"Name_OriginatingCentres_{lang}"
    sc_name_lang = f"Name_SubCentres_{lang}"
    out_headers = make_out_headers(en_h, lang)

    # Build lookups
    oc_lookup = {}  # oc_code -> name (accumulated from multi-row entries)
    sc_lookup = {}  # (oc_code, sc_code) -> name
    no_subcentre_text = ""
    region_texts = {}  # region roman numeral -> text

    cur_oc = None
    cur_oc_parts = []

    for row in ext_rows:
        oc = str(row.get(oc_code, "")).strip()
        sc = str(row.get(sc_code, "")).strip()
        oc_name = str(row.get(oc_name_lang, "")).strip()
        sc_name = str(row.get(sc_name_lang, "")).strip()

        if is_junk(f"{oc} {sc}"):
            continue

        # Region headers
        region_m = re.match(r'^(Р\s*ЕГИОН|РЕГИОН|RÉGION|REGION|REGIÓN)\s*(.+)$', oc_name, re.IGNORECASE)
        if region_m:
            roman = region_m.group(2).strip()
            region_texts[roman] = ndash(oc_name.replace("Р ЕГИОН", "РЕГИОН"))
            cur_oc = None
            continue

        # No sub-centre row
        if sc == "0" and sc_name and any(w in sc_name.lower() for w in ["подцентра", "sub-centre", "centre secondaire", "centro secundario", "no existe"]):
            no_subcentre_text = ndash(sc_name)
            continue

        # Track OC code/name
        if oc and re.match(r'^\d+$', oc):
            if cur_oc != oc:
                cur_oc = oc
                cur_oc_parts = []
            if oc_name:
                cur_oc_parts.append(oc_name)
            oc_lookup[oc] = ndash(" ".join(cur_oc_parts))
        elif not oc and cur_oc and oc_name:
            cur_oc_parts.append(oc_name)
            oc_lookup[cur_oc] = ndash(" ".join(cur_oc_parts))

        # Track SC code/name
        if sc and re.match(r'^\d+$', sc):
            actual_oc = oc if oc and re.match(r'^\d+$', oc) else cur_oc
            if actual_oc:
                sc_lookup[(actual_oc, sc)] = ndash(sc_name)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_id = str(en_row.get("ID", "")).strip()
        en_oc = str(en_row.get(oc_code, "")).strip()
        en_sc = str(en_row.get(sc_code, "")).strip()
        en_oc_name = str(en_row.get("Name_OriginatingCentres_en", "")).strip()
        en_sc_name = str(en_row.get("Name_SubCentres_en", "")).strip()

        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        if "//" in en_id:
            # Special rows
            if en_sc == "0" and "sub-centre" in en_sc_name.lower():
                out[oc_name_lang] = ""
                out[sc_name_lang] = no_subcentre_text
                out["translation_source"] = "PDF" if no_subcentre_text else "NA"
                if no_subcentre_text:
                    matched += 1
                else:
                    unmatched += 1
            elif "REGION" in en_oc_name.upper():
                roman_m = re.search(r'(I+|IV|V|VI)$', en_oc_name)
                found = ""
                if roman_m:
                    found = region_texts.get(roman_m.group(1), "")
                out[oc_name_lang] = found
                out[sc_name_lang] = ""
                out["translation_source"] = "PDF" if found else "NA"
                if found:
                    matched += 1
                else:
                    unmatched += 1
            else:
                out[oc_name_lang] = ""
                out[sc_name_lang] = ""
                out["translation_source"] = "NA"
                unmatched += 1
        else:
            oc_text = oc_lookup.get(en_oc, "")
            sc_text = sc_lookup.get((en_oc, en_sc), "")
            if oc_text or sc_text:
                out[oc_name_lang] = oc_text
                out[sc_name_lang] = sc_text
                out["translation_source"] = "PDF"
                matched += 1
            else:
                out[oc_name_lang] = ""
                out[sc_name_lang] = ""
                out["translation_source"] = "NA"
                unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c13(lang):
    """C13: compound key (DataCategories + SubCategories codes)."""
    en_h, en_rows = read_csv(EN_DIR / "C13_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C13_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    dc_code = "CodeFigure_DataCategories"
    sc_code = "CodeFigure_InternationalDataSubcategories"
    dc_name_lang = f"Name_DataCategories_{lang}"
    sc_name_lang = f"Name_InternationalDataSubcategories_{lang}"
    out_headers = make_out_headers(en_h, lang)

    dc_lookup = {}
    sc_lookup = {}
    cur_dc = None

    for row in ext_rows:
        dc = str(row.get(dc_code, "")).strip()
        sc = str(row.get(sc_code, "")).strip()
        dc_name = str(row.get(dc_name_lang, "")).strip()
        sc_name = str(row.get(sc_name_lang, "")).strip()

        if is_junk(f"{dc} {sc}"):
            continue

        if dc and re.match(r'^\d+$', dc):
            cur_dc = dc
            if dc_name:
                dc_lookup[dc] = ndash(dc_name)

        if sc and re.match(r'^\d+$', sc):
            actual_dc = dc if dc and re.match(r'^\d+$', dc) else cur_dc
            if actual_dc:
                sc_lookup[(actual_dc, sc)] = ndash(sc_name)

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_dc = str(en_row.get(dc_code, "")).strip()
        en_sc = str(en_row.get(sc_code, "")).strip()

        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")

        dc_text = dc_lookup.get(en_dc, "")
        sc_text = sc_lookup.get((en_dc, en_sc), "")

        if dc_text or sc_text:
            out[dc_name_lang] = dc_text
            out[sc_name_lang] = sc_text
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[dc_name_lang] = ""
            out[sc_name_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


def align_c14(lang):
    """C14: CodeFigure (numeric). Fix chemical formulas."""
    en_h, en_rows = read_csv(EN_DIR / "C14_en.csv")
    ext_path = BASE_DIR / lang / "cct_output" / f"C14_{lang}.csv"
    _, ext_rows = read_csv(ext_path)

    text_col_lang = f"Meaning_{lang}"
    out_headers = make_out_headers(en_h, lang)

    lookup = {}
    for row in ext_rows:
        code = str(row.get("CodeFigure", "")).strip()
        text = str(row.get(text_col_lang, "")).strip()
        formula = str(row.get("ChemicalFormula", "")).strip()

        if is_junk(code):
            continue
        if re.match(r'^\d+$', code):
            lookup[code] = (ndash(text), fix_formula(formula))

    out_rows = []
    matched = unmatched = 0

    for en_row in en_rows:
        en_code = str(en_row.get("CodeFigure", "")).strip()
        out = {}
        for h_en, h_out in zip(en_h, out_headers):
            if h_en.endswith("_en"):
                continue
            if h_en == "translation_source":
                continue
            out[h_out] = en_row.get(h_en, "")
        out["Status"] = en_row.get("Status", "Operational")
        # Always use English formula
        out["ChemicalFormula"] = en_row.get("ChemicalFormula", "")

        data = lookup.get(en_code)
        if data:
            out[text_col_lang] = data[0]
            out["translation_source"] = "PDF"
            matched += 1
        else:
            out[text_col_lang] = ""
            out["translation_source"] = "NA"
            unmatched += 1

        out_rows.append(out)

    write_csv(ext_path, out_headers, out_rows)
    return len(en_rows), matched, unmatched


# ── Main ─────────────────────────────────────────────────────────────────────

HANDLERS = {
    "C00": align_c00, "C01": align_c01, "C02": align_c02, "C03": align_c03,
    "C04": align_c04, "C05": align_c05, "C06": align_c06, "C07": align_c07,
    "C08": align_c08, "C11": align_c11, "C12": align_c12, "C13": align_c13,
    "C14": align_c14,
}


def main():
    results = []

    for lang in LANGUAGES:
        print(f"\n{'='*60}")
        print(f"  Language: {lang.upper()}")
        print(f"{'='*60}")

        for table in TABLES:
            try:
                handler = HANDLERS[table]
                total, matched, unmatched = handler(lang)
                pct = (matched / total * 100) if total > 0 else 0
                tag = "OK" if pct > 80 else "LOW" if pct > 50 else "WARN"
                print(f"  {table}_{lang}: {matched}/{total} ({pct:.1f}%) [{tag}]")
                results.append((table, lang, total, matched, unmatched, pct))
            except Exception as e:
                print(f"  {table}_{lang}: ERROR - {e}")
                import traceback; traceback.print_exc()
                results.append((table, lang, 0, 0, 0, 0))

    # Summary
    total_rows = sum(r[2] for r in results)
    total_matched = sum(r[3] for r in results)
    print(f"\n{'='*60}")
    print(f"  SUMMARY: {total_matched}/{total_rows} ({total_matched/total_rows*100:.1f}%)")
    print(f"{'='*60}")

    below = [(r[0], r[1], r[3], r[2], r[5]) for r in results if r[5] < 70]
    if below:
        print("  Below 70%:")
        for t, l, m, tot, p in below:
            print(f"    {t}_{l}: {m}/{tot} ({p:.1f}%)")


if __name__ == "__main__":
    main()
