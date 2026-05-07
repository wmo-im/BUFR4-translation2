"""Extract BUFR Tables A, B, C, and CodeFlag from PDF files.

All extraction is position-based using pymupdf (fitz):
  - ``page.get_text("dict")`` returns spans with bbox coordinates.
  - Spans are classified into columns by x-position (using LangConfig.col_bounds_*).
  - Row boundaries come from FXY positions (y-coordinate grouping).

Table A is the exception: it uses pymupdf4llm markdown output because its
simple single-column structure doesn't have multi-line cell alignment issues.

Public functions
----------------
  extract_table_a(pdf_path, start_page, end_page, lang_config) → pd.DataFrame
  extract_table_b(pdf_path, start_page, end_page, lang, lang_config) → dict[class_no, DataFrame]
  extract_table_c(pdf_path, start_page, end_page, lang, lang_config) → pd.DataFrame
  extract_codeflag(pdf_path, start_page, end_page, lang, lang_config) → dict[class_no, DataFrame]
"""

from __future__ import annotations

import re
from typing import Optional
import fitz
import pandas as pd

from lang_config import LangConfig


# ── Shared helpers ────────────────────────────────────────────────────────────

def _lang_col(base: str, lang: str) -> str:
    return f"{base}_{lang}"


def _get_spans(page) -> list[dict]:
    """Extract all non-empty text spans with bbox from a page."""
    blocks = page.get_text("dict")
    spans = []
    for block in blocks["blocks"]:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if text:
                    spans.append({
                        "text": text,
                        "x0": span["bbox"][0],
                        "y0": span["bbox"][1],
                        "x1": span["bbox"][2],
                        "y1": span["bbox"][3],
                    })
    return spans


def _classify_column(x0: float, col_bounds: dict) -> str | None:
    """Return column name for *x0*, or None if outside all bounds."""
    for col_name, (lo, hi) in col_bounds.items():
        if lo <= x0 < hi:
            return col_name
    return None


def _parse_fxy(f: str, x: str, y: str) -> str:
    """Combine F, X, Y into a 6-digit FXY string (e.g. '001003')."""
    return f"{int(f):01d}{int(x):02d}{int(y):03d}"


def _clean_text(text: str) -> str:
    """Normalize common PDF extraction artifacts."""
    text = re.sub(r"(\w)- (\w)", r"\1-\2", text)   # fix line-break hyphens
    text = re.sub(r" (–\d)", r"\1", text)            # fix superscript spacing
    text = re.sub(r"  +", " ", text)                 # collapse double spaces
    text = text.replace("\x08", "").replace("\x07", "")
    return text.strip()


def _clean_numeric(text: str) -> str:
    """Normalize numeric column values."""
    text = text.strip()
    # Strip continuation artifacts that can leak into numeric cells
    text = re.sub(r"\s*\(à suivre\).*", "", text)
    text = re.sub(r"\s*\(продолж\.\).*", "", text)
    text = re.sub(r"\s*\(continued\).*", "", text)
    text = re.sub(r"\s*\(continúa\).*", "", text)
    text = text.replace("\u2013", "-")  # en-dash → hyphen for negative numbers
    text = text.replace("\x08", "").replace("\x07", "")
    return text.strip()


# ── Table A extraction ────────────────────────────────────────────────────────
# Table A has a simple "CodeFigure  Meaning" row structure — no multi-column
# alignment needed. pymupdf4llm markdown is sufficient.

def _find_first_marker(md: str, markers: list[str], start: int = 0) -> int:
    md_cf = md.casefold()
    found = -1
    for marker in markers:
        pos = md_cf.find(marker.casefold(), start)
        if pos != -1 and (found == -1 or pos < found):
            found = pos
    return found


def extract_table_a(
    pdf_path: str,
    start_page: int,
    end_page: int,
    lang: str,
    lang_config: LangConfig,
) -> pd.DataFrame:
    """Extract Table A (data categories) from PDF pages (0-indexed).

    Uses pymupdf4llm markdown output — Table A's simple structure makes
    position-based extraction unnecessary.

    Returns DataFrame with columns: Id, CodeFigure, Meaning_{lang}, Status
    """
    try:
        import pymupdf4llm
    except ImportError:
        raise ImportError("pymupdf4llm is required for Table A extraction. "
                          "Install with: pip install pymupdf4llm")

    md = pymupdf4llm.to_markdown(pdf_path, pages=list(range(start_page, end_page + 1)))

    # Find the Table A section within the markdown
    start = _find_first_marker(md, lang_config.table_a_start_markers)
    if start == -1:
        raise ValueError(
            f"Could not find Table A start marker in pages {start_page}-{end_page}. "
            f"Check lang_config.table_a_start_markers for '{lang_config.code}'."
        )

    end = len(md)
    for marker in lang_config.table_a_end_markers:
        pos = _find_first_marker(md, [marker], start + 50)
        if pos != -1 and pos < end:
            end = pos

    section = md[start:end]

    # Parse "CodeFigure  Meaning" lines
    meaning_col = _lang_col("Meaning", lang)
    pattern = re.compile(r"^(\d+(?:[–\-]\d+)?)\s+(.+)$")
    rows = []
    for line in section.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = pattern.match(line)
        if m:
            code_figure = m.group(1).replace("–", "-")
            meaning = m.group(2).strip()
            rows.append({
                "Id": f"bufr4/a/{code_figure}",
                "CodeFigure": code_figure,
                meaning_col: meaning,
                "Status": "Operational",
            })

    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ── Table B helpers ───────────────────────────────────────────────────────────

# Tab-separated FXY+name span patterns (French/Spanish/English PDFs)
_TAB_FXY_PATTERN = re.compile(
    r"^(\d)\s*\t\s*(\d{2})\s*\t\s*(\d{3})\s*\t\s*(.+?)(?:\s*\t\s*(.+))?$"
)
_TAB_FXY_ONLY_PATTERN = re.compile(r"^(\d)\s*\t\s*(\d{2})\s*\t\s*(\d{3})\s*$")
_SPACE_FXY_PATTERN = re.compile(r"^(\d)\s+(\d{2})\s+(\d{3})\s+(.+)$")
_SPACE_FXY_ONLY_PATTERN = re.compile(r"^(\d)\s+(\d{2})\s+(\d{3})$")

# Known unit strings — longest first to avoid partial matches.
_KNOWN_UNITS = sorted([
    "Numérique", "Numeric", "Числовой",
    "CCITT IA5", "Table de code", "Table d'indicateurs",
    "Code table", "Flag table", "Кодовая таблица", "Таблица флагов",
    "Tabla de cifrado", "Tabla de banderines", "Numérica", "Carácter",
    "K", "Pa", "m", "s", "kg", "m/s", "m s", "deg", "Hz",
    "Caractère", "Character", "Символьный",
], key=len, reverse=True)


def _emit_fxy(result: list, s: dict, f: str, x: str, y: str) -> None:
    base_x = s["x0"]
    result.append({"text": f, "x0": base_x, "y0": s["y0"], "x1": base_x + 8, "y1": s["y1"]})
    result.append({"text": x, "x0": base_x + 15, "y0": s["y0"], "x1": base_x + 30, "y1": s["y1"]})
    result.append({"text": y, "x0": base_x + 35, "y0": s["y0"], "x1": base_x + 50, "y1": s["y1"]})


def _emit_fxy_name(result: list, s: dict, f: str, x: str, y: str,
                   name_text: str | None, unit_text: str | None,
                   bufr_unit_lo: float) -> None:
    _emit_fxy(result, s, f, x, y)
    if name_text:
        result.append({"text": name_text.strip(), "x0": 150.0,
                       "y0": s["y0"], "x1": 300.0, "y1": s["y1"]})
    if unit_text:
        result.append({"text": unit_text.strip(), "x0": bufr_unit_lo + 1,
                       "y0": s["y0"], "x1": bufr_unit_lo + 50, "y1": s["y1"]})


def _split_merged_spans(spans: list[dict], col_bounds: dict) -> list[dict]:
    """Normalize tab/space-merged FXY+name spans into individual virtual spans.

    French and Spanish PDFs have three FXY span formats in Table B:
    1. Tab-separated with name: "0\\t 01\\t 001\\t Element name"
    2. Tab-separated FXY only:  "0\\t 01\\t 152"
    3. Space-separated:         "0 01 001 Element name"

    All three are split into separate F / XX / YYY / name / unit virtual spans
    so the column-classification logic sees them identically.
    """
    result = []
    fxy_hi = col_bounds.get("fxy", (85, 150))[1]
    bufr_unit_lo = col_bounds.get("bufr_unit", (310, 400))[0]

    for s in spans:
        if "\t" in s["text"]:
            m = _TAB_FXY_PATTERN.match(s["text"])
            if m:
                _emit_fxy_name(result, s, m.group(1), m.group(2), m.group(3),
                               m.group(4), m.group(5), bufr_unit_lo)
                continue
            m_only = _TAB_FXY_ONLY_PATTERN.match(s["text"])
            if m_only:
                _emit_fxy(result, s, m_only.group(1), m_only.group(2), m_only.group(3))
                continue
            result.append(s)
            continue

        if s["x0"] < fxy_hi:
            m_sp = _SPACE_FXY_PATTERN.match(s["text"])
            if m_sp:
                name_text = m_sp.group(4)
                unit_text = None
                for u in _KNOWN_UNITS:
                    if name_text.endswith(" " + u):
                        unit_text = u
                        name_text = name_text[: -len(u)].strip()
                        break
                _emit_fxy_name(result, s, m_sp.group(1), m_sp.group(2), m_sp.group(3),
                               name_text, unit_text, bufr_unit_lo)
                continue
            m_sp_only = _SPACE_FXY_ONLY_PATTERN.match(s["text"])
            if m_sp_only:
                _emit_fxy(result, s, m_sp_only.group(1), m_sp_only.group(2), m_sp_only.group(3))
                continue

        result.append(s)

    return result


def _extract_table_b_page(page, note_pattern: re.Pattern, lang: str,
                           col_bounds: dict, tab_separated: bool) -> list[dict]:
    """Extract Table B rows from one page. Returns list of row dicts."""
    spans = _get_spans(page)
    if tab_separated:
        spans = _split_merged_spans(spans, col_bounds)

    col_spans: dict[str, list] = {col: [] for col in col_bounds}
    for s in spans:
        col = _classify_column(s["x0"], col_bounds)
        if col:
            col_spans[col].append(s)
    for col in col_spans:
        col_spans[col].sort(key=lambda s: s["y0"])

    # Group FXY spans into triples (F, XX, YYY) at the same y
    fxy_spans = col_spans["fxy"]
    fxy_groups: list[list] = []
    current_group: list = []
    current_y = None
    for s in fxy_spans:
        if current_y is None or abs(s["y0"] - current_y) < 3:
            current_group.append(s)
            current_y = s["y0"] if current_y is None else current_y
        else:
            if len(current_group) == 3:
                fxy_groups.append(current_group)
            current_group = [s]
            current_y = s["y0"]
    if len(current_group) == 3:
        fxy_groups.append(current_group)

    # Build row-start list from valid FXY triples
    row_starts = []
    for group in fxy_groups:
        group.sort(key=lambda s: s["x0"])
        f, x, y = group[0]["text"], group[1]["text"], group[2]["text"]
        if not (f.isdigit() and x.isdigit() and y.isdigit()):
            continue
        fxy_str = _parse_fxy(f, x, y)
        row_starts.append({"fxy": fxy_str, "y": group[0]["y0"]})

    if not row_starts:
        return []

    page_bottom = page.rect.height
    for i, row in enumerate(row_starts):
        row["y_end"] = row_starts[i + 1]["y"] - 1 if i + 1 < len(row_starts) else page_bottom

    def _collect(col_name: str, y_start: float, y_end: float) -> str:
        texts = [s["text"] for s in col_spans.get(col_name, [])
                 if y_start - 1 <= s["y0"] <= y_end]
        return " ".join(texts)

    rows = []
    for ri in row_starts:
        y0, y1, fxy = ri["y"], ri["y_end"], ri["fxy"]
        name = _clean_text(_collect("name", y0, y1))
        note = ""
        note_match = note_pattern.search(name)
        if note_match:
            note = note_match.group(0)
            name = (name[:note_match.start()] + name[note_match.end():]).strip().rstrip(",").strip()

        rows.append({
            "FXY": fxy,
            _lang_col("ElementName", lang): name,
            _lang_col("BUFR_Unit", lang): _clean_text(_collect("bufr_unit", y0, y1)),
            "BUFR_Scale": _clean_numeric(_collect("bufr_scale", y0, y1)),
            "BUFR_ReferenceValue": _clean_numeric(_collect("bufr_ref", y0, y1)),
            "BUFR_DataWidth_Bits": _clean_numeric(_collect("bufr_width", y0, y1)),
            _lang_col("CREX_Unit", lang): _clean_text(_collect("crex_unit", y0, y1)),
            "CREX_Scale": _clean_numeric(_collect("crex_scale", y0, y1)),
            "CREX_DataWidth_Char": _clean_numeric(_collect("crex_width", y0, y1)),
            _lang_col("Note", lang): note,
            "Status": "Operational",
        })
    return rows


def _detect_classes(
    pdf_path: str, start_page: int, end_page: int,
    class_pattern: re.Pattern, continuation_filter: str,
) -> list[dict]:
    """Detect Table B class boundaries from page text."""
    doc = fitz.open(pdf_path)
    raw = []
    for page_num in range(start_page, end_page + 1):
        text = doc[page_num].get_text()
        for m in class_pattern.finditer(text):
            name = m.group(2).strip()
            if continuation_filter not in name:
                raw.append({"class_no": m.group(1).zfill(2), "name": name, "page": page_num})
    doc.close()

    seen: set[str] = set()
    classes = []
    for c in raw:
        if c["class_no"] not in seen:
            seen.add(c["class_no"])
            classes.append(c)

    for i, c in enumerate(classes):
        c["start_page"] = c["page"]
        c["end_page"] = classes[i + 1]["page"] - 1 if i + 1 < len(classes) else end_page
    return classes


def extract_table_b(
    pdf_path: str,
    start_page: int,
    end_page: int,
    lang: str,
    lang_config: LangConfig,
) -> dict[str, pd.DataFrame]:
    """Extract all Table B classes from PDF pages (0-indexed).

    Returns dict mapping class_no → DataFrame with columns:
    Id, ClassNo, ClassName_{lang}, FXY, ElementName_{lang},
    BUFR_Unit_{lang}, BUFR_Scale, BUFR_ReferenceValue, BUFR_DataWidth_Bits,
    CREX_Unit_{lang}, CREX_Scale, CREX_DataWidth_Char, Note_{lang}, noteIDs, Status
    """
    classes = _detect_classes(pdf_path, start_page, end_page,
                               lang_config.class_pattern, lang_config.continuation_filter)
    doc = fitz.open(pdf_path)
    results: dict[str, pd.DataFrame] = {}

    for cls in classes:
        all_rows: list[dict] = []
        for page_num in range(cls["start_page"], cls["end_page"] + 1):
            rows = _extract_table_b_page(
                doc[page_num], lang_config.note_pattern_b, lang,
                lang_config.col_bounds_b, lang_config.tab_separated_b,
            )
            all_rows.extend(rows)

        if not all_rows:
            continue

        df = pd.DataFrame(all_rows)
        df = df[df["FXY"].str[1:3] == cls["class_no"]].copy()
        if df.empty:
            continue

        class_name_col = _lang_col("ClassName", lang)
        if lang_config.class_name_prefix:
            display_name = f"{lang_config.class_name_prefix} {cls['class_no']} \u2014 {cls['name']}"
        else:
            display_name = cls["name"]

        df["ClassNo"] = cls["class_no"]
        df[class_name_col] = display_name
        df["Id"] = df["FXY"].apply(lambda fxy: f"bufr4/{fxy[0]}/{fxy[1:3]}/{fxy[3:]}")
        df["noteIDs"] = ""

        cols = ["Id", "ClassNo", class_name_col, "FXY",
                _lang_col("ElementName", lang), _lang_col("BUFR_Unit", lang),
                "BUFR_Scale", "BUFR_ReferenceValue", "BUFR_DataWidth_Bits",
                _lang_col("CREX_Unit", lang), "CREX_Scale", "CREX_DataWidth_Char",
                _lang_col("Note", lang), "noteIDs", "Status"]
        df = df[[c for c in cols if c in df.columns]]
        results[cls["class_no"]] = df

    doc.close()
    return results


# ── Table C extraction ────────────────────────────────────────────────────────

_MERGED_FX_PATTERN = re.compile(r"^(\d)\s+(\d{2})$")


def _split_fx_spans(spans: list[dict]) -> list[dict]:
    """Split merged 'F  XX' spans (e.g. '2  01') into separate virtual spans."""
    result = []
    for s in spans:
        m = _MERGED_FX_PATTERN.match(s["text"])
        if m:
            result.append({"text": m.group(1), "x0": s["x0"], "y0": s["y0"]})
            result.append({"text": m.group(2), "x0": s["x0"] + 15, "y0": s["y0"]})
        else:
            result.append(s)
    return result


def _clean_text_c(text: str) -> str:
    """Normalize Table C text — extra footer patterns beyond shared _clean_text."""
    text = text.replace("\x08", "").replace("\x07", "")
    text = text.replace("\t", " ")
    text = re.sub(r"(\w)- (\w)", r"\1-\2", text)
    text = re.sub(r" (–\d)", r"\1", text)
    text = text.replace("\u00ad ", "").replace("\u00ad", "")
    text = re.sub(r"  +", " ", text)
    text = re.sub(r"\s*\(à suivre\)\s*I\.2\s*–\s*Table\s+C\s+BUFR\s*—\s*\d+\s*$", "", text)
    text = re.sub(r"\s*\(continued\)\s*I\.2\s*–\s*BUFR\s+Table\s+C\s*—\s*\d+\s*$", "", text)
    text = re.sub(r"\s*\(continued\)\s*$", "", text)
    return text.strip()


def _extract_table_c_page(page, note_pattern: re.Pattern, lang: str,
                           col_bounds: dict) -> list[dict]:
    """Extract Table C rows from one page."""
    spans = _get_spans(page)

    col_spans: dict[str, list] = {col: [] for col in col_bounds}
    for s in spans:
        col = _classify_column(s["x0"], col_bounds)
        if col:
            col_spans[col].append(s)
    for col in col_spans:
        col_spans[col].sort(key=lambda s: s["y0"])

    col_spans["fx"] = _split_fx_spans(col_spans["fx"])

    # Group FX spans into pairs (F, XX)
    fx_spans = col_spans["fx"]
    fx_groups: list[list] = []
    current_group: list = []
    current_y = None
    for s in fx_spans:
        if current_y is None or abs(s["y0"] - current_y) < 3:
            current_group.append(s)
            current_y = s["y0"] if current_y is None else current_y
        else:
            if len(current_group) == 2:
                fx_groups.append(current_group)
            current_group = [s]
            current_y = s["y0"]
    if len(current_group) == 2:
        fx_groups.append(current_group)

    row_starts = []
    for group in fx_groups:
        group.sort(key=lambda s: s["x0"])
        f_text, x_text = group[0]["text"], group[1]["text"]
        if not (f_text.isdigit() and x_text.isdigit()):
            continue
        row_starts.append({"f": f_text, "x": x_text, "y": group[0]["y0"]})

    if not row_starts:
        return []

    page_bottom = page.rect.height
    for i, row in enumerate(row_starts):
        row["y_end"] = row_starts[i + 1]["y"] - 1 if i + 1 < len(row_starts) else page_bottom

    def _collect(col_name: str, y_start: float, y_end: float) -> str:
        texts = [s["text"] for s in col_spans[col_name]
                 if y_start - 1 <= s["y0"] <= y_end]
        return " ".join(texts)

    rows = []
    for ri in row_starts:
        y0, y1 = ri["y"], ri["y_end"]
        operand = _collect("operand", y0, y1).strip()
        name = _clean_text_c(_collect("name", y0, y1))
        definition = _clean_text_c(_collect("definition", y0, y1))

        fxy = f"{ri['f']}{ri['x'].zfill(2)}{operand}"

        note = ""
        note_match = note_pattern.search(definition)
        if note_match:
            note = note_match.group(0)

        rows.append({
            "FXY": fxy,
            _lang_col("OperatorName", lang): name,
            _lang_col("OperationDefinition", lang): definition,
            _lang_col("Note", lang): note,
            "Status": "Operational",
        })
    return rows


def extract_table_c(
    pdf_path: str,
    start_page: int,
    end_page: int,
    lang: str,
    lang_config: LangConfig,
) -> pd.DataFrame:
    """Extract Table C (data description operators) from PDF pages (0-indexed).

    Returns DataFrame with columns:
    Id, FXY, OperatorName_{lang}, OperationDefinition_{lang},
    Note_{lang}, noteIDs, Status
    """
    doc = fitz.open(pdf_path)
    all_rows: list[dict] = []
    for page_num in range(start_page, end_page + 1):
        rows = _extract_table_c_page(doc[page_num], lang_config.note_pattern_c,
                                      lang, lang_config.col_bounds_c)
        all_rows.extend(rows)
    doc.close()

    if not all_rows:
        return pd.DataFrame()

    df = pd.DataFrame(all_rows)
    df = df[df["FXY"].str.startswith("2")].copy()
    if df.empty:
        return pd.DataFrame()

    name_col = _lang_col("OperatorName", lang)
    def_col = _lang_col("OperationDefinition", lang)
    for col in [name_col, def_col]:
        if col in df.columns:
            df[col] = df[col].str.replace(r"(\w)- (\w)", r"\1\2", regex=True)

    df["Id"] = df["FXY"].apply(lambda fxy: f"bufr4/{fxy}")
    df["noteIDs"] = ""

    cols = ["Id", "FXY", name_col, def_col, _lang_col("Note", lang), "noteIDs", "Status"]
    return df[[c for c in cols if c in df.columns]]


# ── CodeFlag extraction ───────────────────────────────────────────────────────

FXY_PATTERN = re.compile(r"^(\d)\s+(\d{2})\s+(\d{3})$")
CODE_FIG_PATTERN = re.compile(
    r"^(\d+(?:[–\-]\d+)?)$"                        # numeric or numeric range
    r"|^(?:tou(?:s|tes?)|all|todo[s]?|los|все)\s+\d+$",  # All-N variants (all languages)
    re.IGNORECASE,
)
_FOOTER_REMNANT = re.compile(r"\s+I\.\s*(?:\d+)?\s*\x08?\s*$")

# French missing-value rows: "4 bits mis à 1" (or with "Valeur manquante" embedded)
# get incorrectly merged into the previous entry's name because the code-figure
# "4 bits mis à 1" doesn't match CODE_FIG_PATTERN.  Detected in post-processing.
_FR_ALLBITS = re.compile(
    r"\s+(\d+)\s+bits?\s+(?:Valeur\s+manquante\s+)?mis\s+à\s+1\b.*$",
    re.IGNORECASE,
)

# Inline note-text patterns (strip from EntryName; full text goes to notes files)
_NOTE_PATTERNS: dict[str, re.Pattern] = {
    "fr": re.compile(r"\s+NOTES?\s*[:\t]", re.IGNORECASE),
    "ru": re.compile(r"\s+Примечани[ея]\.?\s*:?\s*"),
    "es": re.compile(r"\s+Notas?\s*[.:\t]", re.IGNORECASE),
    "en": re.compile(r"\s+NOTES?\s*[:\t]", re.IGNORECASE),
}

# Reference-note patterns per language
_REF_NOTE_PATTERNS: dict[str, re.Pattern] = {
    "ru": re.compile(r"\(См\.\s+общую кодовую таблицу\s+C[–\-]\d+.*?\)"),
    "fr": re.compile(r"\(Voir\s+table\s+de\s+code\s+commune\s+C[–\-]\d+.*?\)", re.IGNORECASE),
    "en": re.compile(r"\(See\s+common\s+Code\s+table\s+C[–\-]\d+.*?\)", re.IGNORECASE),
    "es": re.compile(r"\(V[eé]ase\s+tabla\s+de\s+c[oó]digo\s+com[uú]n\s+C[–\-]\d+.*?\)", re.IGNORECASE),
}


def _clean_entry_text(text: str) -> str:
    text = text.replace("\x08", "").replace("\x07", "")
    text = _FOOTER_REMNANT.sub("", text)
    text = re.sub(r" (–\d)", r"\1", text)
    return text.strip()


def _is_fxy_header(span: dict, x_fxy_center: float) -> tuple[str, str, str] | None:
    """Return (F, XX, YYY) if span is a centered FXY header, else None."""
    if span["x0"] < x_fxy_center:
        return None
    m = FXY_PATTERN.match(span["text"])
    if m:
        return m.group(1), m.group(2), m.group(3)
    return None


def _is_marker(text: str, x0: float, marker_phrases: frozenset,
               x_code_fig_max: float) -> bool:
    return bool((text in marker_phrases or text.strip() in marker_phrases)
                and x0 < x_code_fig_max + 50)


def _is_page_header_footer(span: dict, page_height: float,
                           footer_texts: frozenset, footer_keyword: str) -> bool:
    text = span["text"].strip()
    if span["y0"] < 50:
        return True
    if span["y0"] > page_height - 90:
        return True
    if text in footer_texts:
        return True
    if footer_keyword and footer_keyword in text:
        return True
    if span["y0"] > page_height - 70 and text.startswith("—"):
        return True
    return False


def _extract_codeflag_page(page, config: LangConfig, lang: str,
                            previous_fxy: str | None) -> dict:
    """Extract CodeFlag entries from one page.

    Returns dict:
        sections: list of {fxy, element_name, entries, is_continuation}
        continuation_fxy: str | None
    """
    spans = _get_spans(page)
    page_height = page.rect.height
    x_fxy_center = config.codeflag_x_fxy_center
    x_code_fig_max = config.codeflag_x_code_fig_max
    x_entry_name_min = config.codeflag_x_entry_name_min
    x_sub1_min = config.codeflag_x_sub1_min
    marker_phrases = config.codeflag_marker_phrases
    footer_texts = config.codeflag_footer_texts
    footer_keyword = config.codeflag_footer_keyword
    continuation_pattern = config.codeflag_continuation_pattern
    ref_note_pattern = _REF_NOTE_PATTERNS.get(lang, _REF_NOTE_PATTERNS["en"])

    continuation_fxy = None
    if continuation_pattern:
        page_text = page.get_text()
        cont_match = continuation_pattern.search(page_text)
        if cont_match:
            fxy_match = re.search(r"(\d)\s+(\d{2})\s+(\d{3})", cont_match.group(0))
            if not fxy_match:
                fxy_match = re.search(r"(\d)\s+(\d{2})\s+(\d{3})", page_text)
            if fxy_match:
                continuation_fxy = _parse_fxy(
                    fxy_match.group(1), fxy_match.group(2), fxy_match.group(3)
                )
            elif previous_fxy:
                continuation_fxy = previous_fxy

    spans = [s for s in spans
             if not _is_page_header_footer(s, page_height, footer_texts, footer_keyword)]
    spans.sort(key=lambda s: (s["y0"], s["x0"]))

    # Phase 1: find FXY headers and continuation markers
    fxy_headers = []
    for s in spans:
        if continuation_pattern and continuation_pattern.search(s["text"]):
            m = re.search(r"(\d)\s+(\d{2})\s+(\d{3})", s["text"])
            if m:
                continuation_fxy = _parse_fxy(m.group(1), m.group(2), m.group(3))
            continue
        fxy_match = _is_fxy_header(s, x_fxy_center)
        if fxy_match:
            fxy_str = _parse_fxy(*fxy_match)
            fxy_headers.append({"fxy": fxy_str, "y": s["y0"]})

    if not fxy_headers and not continuation_fxy and previous_fxy:
        continuation_fxy = previous_fxy

    # Phase 2: define y-ranges per section
    sections = []
    if continuation_fxy:
        y_end = fxy_headers[0]["y"] - 1 if fxy_headers else page_height
        sections.append({"fxy": continuation_fxy, "y_start": 0,
                         "y_end": y_end, "is_continuation": True})

    for i, hdr in enumerate(fxy_headers):
        y_end = fxy_headers[i + 1]["y"] - 1 if i + 1 < len(fxy_headers) else page_height
        sections.append({"fxy": hdr["fxy"], "y_start": hdr["y"], "y_end": y_end})

    # Phase 3: extract entries per section
    result_sections = []
    entry_col = _lang_col("EntryName", lang)
    sub1_col = _lang_col("EntryName_sub1", lang)

    for sec in sections:
        fxy = sec["fxy"]
        y_start, y_end = sec["y_start"], sec["y_end"]
        is_cont = sec.get("is_continuation", False)

        sec_spans = [s for s in spans
                     if y_start - 1 <= s["y0"] <= y_end
                     and not _is_fxy_header(s, x_fxy_center)]

        first_code_y = None
        marker_y = None
        for s in sec_spans:
            if s["x0"] < x_code_fig_max + 20 and _is_marker(s["text"], s["x0"],
                                                              marker_phrases, x_code_fig_max):
                marker_y = s["y0"]
                continue
            if s["x0"] < x_code_fig_max + 20 and CODE_FIG_PATTERN.match(s["text"]):
                first_code_y = s["y0"]
                break

        element_name_parts = []
        reference_note = ""
        if not is_cont:
            name_y_end = marker_y or first_code_y or y_end
            for s in sec_spans:
                if s["y0"] <= y_start:
                    continue
                if s["y0"] >= name_y_end:
                    break
                if _is_marker(s["text"], s["x0"], marker_phrases, x_code_fig_max):
                    continue
                ref_match = ref_note_pattern.search(s["text"])
                if ref_match:
                    reference_note = s["text"]
                    continue
                if s["x0"] > 80:
                    element_name_parts.append(s["text"])

        element_name = _clean_entry_text(
            re.sub(r"(\w)- (\w)", r"\1\2", " ".join(element_name_parts))
        )

        entries = []
        if first_code_y is not None:
            candidate_spans = [
                s for s in sec_spans
                if s["y0"] >= first_code_y - 1
                and not _is_marker(s["text"], s["x0"], marker_phrases, x_code_fig_max)
                and not (continuation_pattern and continuation_pattern.search(s["text"]))
            ]
            candidate_spans.sort(key=lambda s: (s["y0"], s["x0"]))

            # Group into rows
            rows_of_spans: list[list] = []
            current_row: list = []
            current_y = None
            for s in candidate_spans:
                if current_y is None or abs(s["y0"] - current_y) < 2:
                    current_row.append(s)
                    current_y = s["y0"] if current_y is None else current_y
                else:
                    if current_row:
                        rows_of_spans.append(current_row)
                    current_row = [s]
                    current_y = s["y0"]
            if current_row:
                rows_of_spans.append(current_row)

            code_fig_spans: list = []
            entry_name_spans: list = []
            sub1_spans: list = []

            for row in rows_of_spans:
                row.sort(key=lambda s: s["x0"])
                found_code_fig = False
                for s in row:
                    if (not found_code_fig
                            and s["x0"] < x_code_fig_max + 20
                            and CODE_FIG_PATTERN.match(s["text"])):
                        code_fig_spans.append(s)
                        found_code_fig = True
                    elif s["x0"] >= x_sub1_min:
                        sub1_spans.append(s)
                    else:
                        entry_name_spans.append(s)

            for i, cf in enumerate(code_fig_spans):
                cf_y = cf["y0"]
                next_y = code_fig_spans[i + 1]["y0"] if i + 1 < len(code_fig_spans) else y_end

                name_texts = [s["text"] for s in entry_name_spans if cf_y - 1 <= s["y0"] < next_y]
                entry_text = _clean_entry_text(
                    re.sub(r"(\w)- (\w)", r"\1\2", " ".join(name_texts))
                )
                sub1_texts = [s["text"] for s in sub1_spans if cf_y - 1 <= s["y0"] < next_y]
                sub1_text = _clean_entry_text(" ".join(sub1_texts))

                entries.append({
                    "CodeFigure": cf["text"].replace("\u2013", "-"),
                    entry_col: entry_text,
                    sub1_col: sub1_text,
                })

        elif reference_note:
            entries.append({"CodeFigure": "", entry_col: reference_note, sub1_col: ""})

        # ── Strip inline note text from entry names ──
        note_pat = _NOTE_PATTERNS.get(lang)
        if note_pat and entries:
            for entry in entries:
                m = note_pat.search(entry[entry_col])
                if m:
                    entry["_note_text"] = entry[entry_col][m.start():].strip()
                    entry[entry_col] = entry[entry_col][:m.start()].strip()
                if entry.get(sub1_col):
                    m2 = note_pat.search(entry[sub1_col])
                    if m2:
                        note_part = entry[sub1_col][m2.start():].strip()
                        entry["_note_text"] = (
                            entry.get("_note_text", "") + " " + note_part
                        ).strip()
                        entry[sub1_col] = entry[sub1_col][:m2.start()].strip()

        # French All-N post-processing: "N bits [Valeur manquante] mis à 1" gets
        # merged into the last entry's EntryName because the code-figure text doesn't
        # match CODE_FIG_PATTERN.  Split it out as a separate All-N entry.
        if lang == "fr" and entries:
            for idx in range(len(entries)):
                m = _FR_ALLBITS.search(entries[idx][entry_col])
                if m:
                    n = m.group(1)
                    entries[idx][entry_col] = entries[idx][entry_col][:m.start()].strip()
                    entries.append({
                        "CodeFigure": f"{n} bits mis à 1",
                        entry_col: "Valeur manquante",
                        sub1_col: "",
                    })
                    break  # at most one All-N per section

        result_sections.append({
            "fxy": fxy,
            "element_name": element_name,
            "entries": entries,
            "is_continuation": is_cont,
        })

    return {"sections": result_sections, "continuation_fxy": continuation_fxy}


def _detect_codeflag_classes(pdf_path: str, start_page: int, end_page: int,
                              footer_pattern: re.Pattern) -> list[dict]:
    """Detect CodeFlag class boundaries from page footer text."""
    doc = fitz.open(pdf_path)
    class_pages: dict[str, list[int]] = {}
    for page_num in range(start_page, end_page + 1):
        text = doc[page_num].get_text()
        m = footer_pattern.search(text)
        if m:
            cls_no = m.group(1).zfill(2)
            class_pages.setdefault(cls_no, []).append(page_num)
    doc.close()

    return [
        {"class_no": cls, "start_page": min(pages), "end_page": max(pages)}
        for cls, pages in sorted(class_pages.items())
    ]


def extract_codeflag(
    pdf_path: str,
    start_page: int,
    end_page: int,
    lang: str,
    lang_config: LangConfig,
) -> dict[str, pd.DataFrame]:
    """Extract all CodeFlag classes from PDF pages (0-indexed).

    Returns dict mapping class_no → DataFrame with columns:
    Id, FXY, ElementName_{lang}, CodeFigure, EntryName_{lang},
    EntryName_sub1_{lang}, EntryName_sub2_{lang}, Note_{lang}, noteIDs, Status
    """
    footer_pattern = lang_config.codeflag_footer_pattern
    if not footer_pattern:
        raise ValueError(f"No codeflag_footer_pattern configured for lang '{lang_config.code}'")

    classes = _detect_codeflag_classes(pdf_path, start_page, end_page, footer_pattern)
    doc = fitz.open(pdf_path)
    results: dict[str, pd.DataFrame] = {}

    element_col = _lang_col("ElementName", lang)
    entry_col = _lang_col("EntryName", lang)
    sub1_col = _lang_col("EntryName_sub1", lang)
    sub2_col = _lang_col("EntryName_sub2", lang)
    note_col = _lang_col("Note", lang)

    for cls in classes:
        all_rows: list[dict] = []
        current_fxy = None
        current_element_name = None

        for page_num in range(cls["start_page"], cls["end_page"] + 1):
            result = _extract_codeflag_page(doc[page_num], lang_config, lang, current_fxy)
            for section in result["sections"]:
                fxy = section["fxy"]
                if fxy[1:3] != cls["class_no"]:
                    continue
                if section["is_continuation"]:
                    element_name = current_element_name or ""
                else:
                    element_name = section["element_name"]
                    current_fxy = fxy
                    current_element_name = element_name

                F, XX, YYY = fxy[0], fxy[1:3], fxy[3:6]
                if section["entries"]:
                    for entry in section["entries"]:
                        code_fig = entry["CodeFigure"]
                        row_id = f"bufr4/{F}/{XX}/{YYY}/{code_fig}" if code_fig else f"bufr4/{F}/{XX}/{YYY}"
                        all_rows.append({
                            "Id": row_id, "FXY": fxy,
                            element_col: element_name, "CodeFigure": code_fig,
                            entry_col: entry.get(entry_col, ""),
                            sub1_col: entry.get(sub1_col, ""),
                            sub2_col: "", note_col: "", "noteIDs": "", "Status": "Operational",
                        })
                else:
                    all_rows.append({
                        "Id": f"bufr4/{F}/{XX}/{YYY}", "FXY": fxy,
                        element_col: element_name, "CodeFigure": "",
                        entry_col: "", sub1_col: "", sub2_col: "",
                        note_col: "", "noteIDs": "", "Status": "Operational",
                    })

        if not all_rows:
            continue

        cols = ["Id", "FXY", element_col, "CodeFigure", entry_col,
                sub1_col, sub2_col, note_col, "noteIDs", "Status"]
        results[cls["class_no"]] = pd.DataFrame(all_rows, columns=cols)

    doc.close()
    return results


# ── Convenience: save all tables for a given type ────────────────────────────

def save_table_b(results: dict[str, pd.DataFrame], output_dir: str, lang: str) -> int:
    """Write Table B class CSVs to *output_dir*. Returns total rows saved."""
    import os
    os.makedirs(output_dir, exist_ok=True)
    total = 0
    for class_no, df in sorted(results.items()):
        path = os.path.join(output_dir, f"BUFRCREX_TableB_{lang}_{class_no}.csv")
        df.to_csv(path, index=False)
        total += len(df)
        print(f"  Class {class_no}: {len(df)} rows → {path}")
    return total


def save_codeflag(results: dict[str, pd.DataFrame], output_dir: str, lang: str) -> int:
    import os
    os.makedirs(output_dir, exist_ok=True)
    total = 0
    for class_no, df in sorted(results.items()):
        path = os.path.join(output_dir, f"BUFRCREX_CodeFlag_{lang}_{class_no}.csv")
        df.to_csv(path, index=False)
        total += len(df)
        print(f"  Class {class_no}: {len(df)} rows → {path}")
    return total
