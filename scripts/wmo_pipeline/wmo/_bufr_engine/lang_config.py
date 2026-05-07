"""Language-specific configuration for BUFR table extraction.

Usage
-----
    from lang_config import get_lang_config
    cfg = get_lang_config("es")   # or "fr", "en", "ru"

To add a new language:
  1. Define _COL_BOUNDS_B_XX, _COL_BOUNDS_C_XX, _COL_BOUNDS_D_XX dicts.
  2. Add a LangConfig entry in LANG_CONFIGS.
  3. Test extraction on a sample page with the new config.

All x-coordinates are in PDF points (1 pt = 1/72 inch) measured from
page left edge. Values were calibrated by reading page.get_text("dict")
on known pages and inspecting span["bbox"][0] values.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Union
import re


@dataclass(frozen=True)
class LangConfig:
    """Frozen config for one language variant of a BUFR WMO-306 PDF.

    Fields
    ------
    code : str
        Two-letter language code (e.g. "es", "fr", "en", "ru").
    table_a_start_markers : list[str]
        Case-insensitive substrings marking the start of Table A in the PDF text.
    table_a_end_markers : list[str]
        Case-insensitive substrings marking the end of Table A.
    class_pattern : re.Pattern
        Regex for Table B class headings, capturing (class_no, name).
        Example match: "Clase 01 — Surface elements" → ("01", "Surface elements")
    category_pattern : re.Pattern
        Regex for Table D category headings, capturing (category_no, name).
    note_pattern_b : re.Pattern
        Regex for footnote references in Table B element names.
    note_pattern_c : re.Pattern
        Regex for footnote references in Table C definitions.
    col_bounds_b/c/d : dict[str, tuple[float, float]]
        Column x-boundaries per table type. Keys are column names;
        values are (x_min, x_max) in PDF points.
    tab_separated_b : bool
        True if Table B FXY codes and element names are merged in tab-separated
        spans (French/Spanish/English PDFs). False for Russian PDF.
    continuation_filter : str
        Substring to filter out class/category continuation headings
        (e.g. "continúa", "suite", "продолж").
    codeflag_footer_pattern : re.Pattern | None
        Regex matching page footer text that identifies CodeFlag class number.
        Capture group 1 must be the class number string.
    codeflag_continuation_pattern : re.Pattern | None
        Regex matching continuation headers like "(Code table 0 02 001 — continued)".
    codeflag_marker_phrases : frozenset[str]
        Column-header words to skip during CodeFlag entry extraction
        (e.g. {"Code figure", "Bit No."}).
    codeflag_footer_texts : frozenset[str]
        Exact span texts that are page headers/footers (to be filtered out).
    codeflag_footer_keyword : str
        Substring that identifies a footer span when found in span text.
    codeflag_x_fxy_center : float
        Minimum x0 for a centered FXY header span. Centered headers have
        x0 > this value, column-data spans have x0 < this value.
    codeflag_x_code_fig_max : float
        Maximum x0 for a code-figure span (left column).
    codeflag_x_entry_name_min : float
        Minimum x0 for an entry-name span (middle column).
    codeflag_x_sub1_min : float
        Minimum x0 for a sub-entry (sub1) span (right column).
    class_name_prefix : str
        Prefix for ClassName display (e.g. "Clase" → "Clase 01 — Name").
        Empty string means use raw name without prefix.
    category_name_prefix : str
        Prefix for CategoryOfSequences display in Table D.
    """

    code: str
    table_a_start_markers: list[str]
    table_a_end_markers: list[str]
    class_pattern: re.Pattern
    category_pattern: re.Pattern
    note_pattern_b: re.Pattern
    note_pattern_c: re.Pattern
    col_bounds_b: dict[str, tuple[float, float]] = field(default_factory=dict)
    col_bounds_c: dict[str, tuple[float, float]] = field(default_factory=dict)
    col_bounds_d: dict[str, tuple[float, float]] = field(default_factory=dict)
    tab_separated_b: bool = False
    continuation_filter: str = ""
    codeflag_footer_pattern: re.Pattern | None = None
    codeflag_continuation_pattern: re.Pattern | None = None
    codeflag_marker_phrases: frozenset[str] = field(default_factory=frozenset)
    codeflag_footer_texts: frozenset[str] = field(default_factory=frozenset)
    codeflag_footer_keyword: str = ""
    codeflag_x_fxy_center: float = 250.0
    codeflag_x_code_fig_max: float = 143.0
    codeflag_x_entry_name_min: float = 143.0
    codeflag_x_sub1_min: float = 390.0
    class_name_prefix: str = ""
    category_name_prefix: str = ""


def _compile(pattern: Union[str, re.Pattern, None]) -> re.Pattern | None:
    if pattern is None:
        return None
    if isinstance(pattern, re.Pattern):
        return pattern
    return re.compile(pattern, re.IGNORECASE)


def normalize_lang(lang: str) -> str:
    return lang.strip().lower().replace("-", "_")


def lang_suffix(lang: str) -> str:
    """Return the lowercase two-letter lang code used as CSV column suffix."""
    return normalize_lang(lang)


# ── Column boundary dictionaries ──────────────────────────────────────────────
# Calibrated per-PDF by inspecting span["bbox"][0] on known pages.
# Add a new language by creating analogous dicts below.

# Russian (WMO-306 2019)
_COL_BOUNDS_B_RU = {
    "fxy": (85, 148), "name": (148, 345), "bufr_unit": (345, 415),
    "bufr_scale": (415, 475), "bufr_ref": (475, 538), "bufr_width": (538, 580),
    "crex_unit": (580, 655), "crex_scale": (655, 715), "crex_width": (715, 760),
}
_COL_BOUNDS_C_RU = {
    "fx": (85, 130), "operand": (130, 182), "name": (182, 295), "definition": (295, 760),
}
_COL_BOUNDS_D_RU = {
    "fxy1": (80, 148), "fxy2": (148, 205), "name": (205, 448), "description": (448, 760),
}

# French
_COL_BOUNDS_B_FR = {
    "fxy": (85, 150), "name": (150, 310), "bufr_unit": (310, 400),
    "bufr_scale": (400, 430), "bufr_ref": (430, 530), "bufr_width": (530, 570),
    "crex_unit": (570, 645), "crex_scale": (645, 705), "crex_width": (705, 760),
}
_COL_BOUNDS_C_FR = {
    "fx": (75, 140), "operand": (140, 190), "name": (190, 300), "definition": (300, 760),
}
_COL_BOUNDS_D_FR = {
    "fxy1": (70, 155), "fxy2": (155, 245), "name": (245, 440), "description": (440, 760),
}

# English
_COL_BOUNDS_B_EN = {
    "fxy": (80, 143), "name": (143, 306), "bufr_unit": (306, 390),
    "bufr_scale": (390, 440), "bufr_ref": (440, 520), "bufr_width": (520, 565),
    "crex_unit": (565, 650), "crex_scale": (650, 710), "crex_width": (710, 760),
}
_COL_BOUNDS_C_EN = {
    "fx": (85, 145), "operand": (145, 197), "name": (197, 295), "definition": (295, 760),
}
_COL_BOUNDS_D_EN = {
    "fxy1": (75, 155), "fxy2": (155, 212), "name": (212, 435), "description": (435, 760),
}

# Spanish (2021 PDF — slightly wider margins)
_COL_BOUNDS_B_ES = {
    "fxy": (55, 130), "name": (130, 362), "bufr_unit": (362, 443),
    "bufr_scale": (443, 487), "bufr_ref": (487, 555), "bufr_width": (555, 628),
    "crex_unit": (628, 692), "crex_scale": (692, 740), "crex_width": (740, 780),
}
_COL_BOUNDS_C_ES = {
    "fx": (55, 120), "operand": (120, 170), "name": (170, 300), "definition": (300, 760),
}
_COL_BOUNDS_D_ES = {
    "fxy1": (55, 125), "fxy2": (125, 195), "name": (195, 440), "description": (440, 760),
}


# ── Pre-built language configs ────────────────────────────────────────────────

LANG_CONFIGS: dict[str, LangConfig] = {
    "ru": LangConfig(
        code="ru",
        table_a_start_markers=["таблица a кода bufr"],
        table_a_end_markers=[
            "таблица в кодов bufr",
            "таблица b кодов bufr",
            "таблицы кода bufr, относящиеся к разделу 3",
        ],
        class_pattern=_compile(r"класс\s+(\d+)\s*[—\-–]\s*(.+)"),
        category_pattern=_compile(r"категория\s+(\d+)\s*[—\-–]\s*(.+)"),
        note_pattern_b=_compile(r"\(см\.\s*примечани[еяй]\s*[\d,\s]+\)"),
        note_pattern_c=_compile(r"\(см\.\s*примечани[еяй]\s*[\d,\sи]+\)"),
        col_bounds_b=_COL_BOUNDS_B_RU,
        col_bounds_c=_COL_BOUNDS_C_RU,
        col_bounds_d=_COL_BOUNDS_D_RU,
        tab_separated_b=False,
        continuation_filter="продолж",
        codeflag_footer_pattern=_compile(r"Кодовые таблицы/Таблицы флагов/(\d+)"),
        codeflag_continuation_pattern=_compile(
            r"\(?(?:Кодовая таблица|Таблица флагов)\s+\d\s+\d{2}\s+\d{3}\s*—\s*продолж\.\)?"
        ),
        codeflag_marker_phrases=frozenset({"Кодовая", "цифра", "Номер бита"}),
        codeflag_footer_texts=frozenset({"FM 94 BUFR", "(продолж.)", "I.2 –", "I", ".2 –"}),
        codeflag_footer_keyword="Кодовые таблицы/Таблицы флагов",
        codeflag_x_fxy_center=250.0,
        codeflag_x_code_fig_max=143.0,
        codeflag_x_entry_name_min=143.0,
        codeflag_x_sub1_min=390.0,
        category_name_prefix="Категория",
    ),
    "fr": LangConfig(
        code="fr",
        table_a_start_markers=[
            "table a du code bufr — catégorie de données",
            "table a du code bufr",
        ],
        table_a_end_markers=[
            "tables du code bufr correspondant à la section 3",
            "tables du code bufr correspondant a la section 3",
            "table b du code bufr/crex",
        ],
        class_pattern=_compile(r"classe\s+(\d+)\s*[—\-–]\s*(.+)"),
        category_pattern=_compile(r"cat[eé]gorie\s+(\d+)\s*[—\-–]\s*(.+)"),
        note_pattern_b=_compile(r"\(voir\s+(?:note|remarque)s?\s*[\d,\s et]+\)"),
        note_pattern_c=_compile(r"\(voir\s+(?:note|remarque)s?\s*[\d,\s et]+\)"),
        col_bounds_b=_COL_BOUNDS_B_FR,
        col_bounds_c=_COL_BOUNDS_C_FR,
        col_bounds_d=_COL_BOUNDS_D_FR,
        tab_separated_b=True,
        continuation_filter="suite",
        codeflag_footer_pattern=_compile(r"Tables/(\d+)\s+CODE/INDICATEURS"),
        codeflag_continuation_pattern=_compile(
            r"\(?(?:Table de code|Table d.indicateurs?)\s+\d\s+\d{2}\s+\d{3}\s*—\s*suite\)?"
        ),
        codeflag_marker_phrases=frozenset({"Chiffre", "du code", "Numéro de bit", "Numéro"}),
        codeflag_footer_texts=frozenset({
            "FM 94 BUFR", "FM 94  BUFR", "(à suivre)", "I.2 –", "I", ".2 –",
            "FM 94 BUFR, FM 95 CREX",
        }),
        codeflag_footer_keyword="CODE/INDICATEURS",
        codeflag_x_fxy_center=250.0,
        codeflag_x_code_fig_max=100.0,
        codeflag_x_entry_name_min=110.0,
        codeflag_x_sub1_min=390.0,
        class_name_prefix="Classe",
        category_name_prefix="",
    ),
    "en": LangConfig(
        code="en",
        table_a_start_markers=["bufr table a"],
        table_a_end_markers=[
            "bufr tables relative to section 3",
            "bufr/crex table b",
            "bufr table a — 1",
        ],
        class_pattern=_compile(r"[Cc]lass\s+(\d+)\s*[—\-–]\s*(.+)"),
        category_pattern=_compile(r"[Cc]ategory\s+(\d+)\s*[—\-–]\s*(.+)"),
        note_pattern_b=_compile(r"\(see\s+Notes?\s*[\d,\s and]+\)"),
        note_pattern_c=_compile(r"\(see\s+Notes?\s*[\d,\s and]+\)"),
        col_bounds_b=_COL_BOUNDS_B_EN,
        col_bounds_c=_COL_BOUNDS_C_EN,
        col_bounds_d=_COL_BOUNDS_D_EN,
        tab_separated_b=True,
        continuation_filter="continued",
        codeflag_footer_pattern=_compile(r"CODE/FLAG Tables/(\d+)"),
        codeflag_continuation_pattern=_compile(
            r"\(?(?:Code table|Flag table)\s+\d\s+\d{2}\s+\d{3}\s*[—\-–]\s*continued\)?"
        ),
        codeflag_marker_phrases=frozenset({"Code figure", "Bit No."}),
        codeflag_footer_texts=frozenset({
            "FM 94 BUFR", "FM 94  BUFR", "(continued)", "I.2 –", "I", ".2 –",
            "FM 94 BUFR, FM 95 CREX",
        }),
        codeflag_footer_keyword="CODE/FLAG Tables",
        codeflag_x_fxy_center=250.0,
        codeflag_x_code_fig_max=110.0,
        codeflag_x_entry_name_min=150.0,
        codeflag_x_sub1_min=390.0,
        category_name_prefix="Category",
    ),
    "es": LangConfig(
        code="es",
        table_a_start_markers=[
            "tabla a bufr",
            "tabla a del codigo bufr",
            "tabla a del código bufr",
            "table a de la clave bufr",
        ],
        table_a_end_markers=[
            "tablas bufr relativas a la sección 3",
            "tablas bufr relativas a la seccion 3",
            "tabla b de las claves bufr/crex",
            "tabla b del codigo bufr/crex",
            "tabla b del código bufr/crex",
            "bufr/crex tabla b",
        ],
        class_pattern=_compile(r"clase\s+(\d+)\s*[—\-–]\s*(.+)"),
        category_pattern=_compile(r"categor[ií]a\s+(\d+)\s*[—\-–]\s*(.+)"),
        note_pattern_b=_compile(r"\(v[ée]ase\s+nota[s]?\s*[\d,\s y]+\)"),
        note_pattern_c=_compile(r"\(v[ée]ase\s+nota[s]?\s*[\d,\s y]+\)"),
        col_bounds_b=_COL_BOUNDS_B_ES,
        col_bounds_c=_COL_BOUNDS_C_ES,
        col_bounds_d=_COL_BOUNDS_D_ES,
        tab_separated_b=True,
        continuation_filter="continúa",
        codeflag_footer_pattern=_compile(r"Tablas de CIFRADO/BANDERINES/(\d+)"),
        codeflag_continuation_pattern=_compile(
            r"\(?(?:Tabla de cifrado|Tabla de banderines)\s+\d\s+\d{2}\s+\d{3}\s*[—\-–]\s*contin[úu]a\)?|\(contin[úu]a\)"
        ),
        codeflag_marker_phrases=frozenset({"Cifra", "de clave", "Número de bit", "Número"}),
        codeflag_footer_texts=frozenset({
            "FM 94 BUFR", "FM 94  BUFR", "(continúa)", "I.2 –", "I", ".2 –",
            "FM 94 BUFR, FM 95 CREX",
        }),
        codeflag_footer_keyword="CIFRADO/BANDERINES",
        codeflag_x_fxy_center=250.0,
        codeflag_x_code_fig_max=110.0,
        codeflag_x_entry_name_min=150.0,
        codeflag_x_sub1_min=390.0,
        class_name_prefix="Clase",
        category_name_prefix="",
    ),
}


def get_lang_config(lang: str, overrides: Optional[dict] = None) -> LangConfig:
    """Return LangConfig for *lang*, with optional field overrides.

    Parameters
    ----------
    lang : str
        Language code — "es", "fr", "en", or "ru".
    overrides : dict, optional
        Dict mapping LangConfig field names to new values.
        Regex-valued fields can be passed as strings (auto-compiled).
    """
    code = normalize_lang(lang)
    if code not in LANG_CONFIGS:
        raise ValueError(
            f"Unsupported language '{lang}'. Supported: {', '.join(LANG_CONFIGS)}"
        )
    base = LANG_CONFIGS[code]
    if not overrides:
        return base

    data = {
        "code": base.code,
        "table_a_start_markers": overrides.get("table_a_start_markers", base.table_a_start_markers),
        "table_a_end_markers": overrides.get("table_a_end_markers", base.table_a_end_markers),
        "class_pattern": _compile(overrides.get("class_pattern", base.class_pattern)),
        "category_pattern": _compile(overrides.get("category_pattern", base.category_pattern)),
        "note_pattern_b": _compile(overrides.get("note_pattern_b", base.note_pattern_b)),
        "note_pattern_c": _compile(overrides.get("note_pattern_c", base.note_pattern_c)),
        "col_bounds_b": overrides.get("col_bounds_b", base.col_bounds_b),
        "col_bounds_c": overrides.get("col_bounds_c", base.col_bounds_c),
        "col_bounds_d": overrides.get("col_bounds_d", base.col_bounds_d),
        "tab_separated_b": overrides.get("tab_separated_b", base.tab_separated_b),
        "continuation_filter": overrides.get("continuation_filter", base.continuation_filter),
        "codeflag_footer_pattern": _compile(
            overrides.get("codeflag_footer_pattern", base.codeflag_footer_pattern)
        ),
        "codeflag_continuation_pattern": _compile(
            overrides.get("codeflag_continuation_pattern", base.codeflag_continuation_pattern)
        ),
        "codeflag_marker_phrases": overrides.get("codeflag_marker_phrases", base.codeflag_marker_phrases),
        "codeflag_footer_texts": overrides.get("codeflag_footer_texts", base.codeflag_footer_texts),
        "codeflag_footer_keyword": overrides.get("codeflag_footer_keyword", base.codeflag_footer_keyword),
        "codeflag_x_fxy_center": overrides.get("codeflag_x_fxy_center", base.codeflag_x_fxy_center),
        "codeflag_x_code_fig_max": overrides.get("codeflag_x_code_fig_max", base.codeflag_x_code_fig_max),
        "codeflag_x_entry_name_min": overrides.get("codeflag_x_entry_name_min", base.codeflag_x_entry_name_min),
        "codeflag_x_sub1_min": overrides.get("codeflag_x_sub1_min", base.codeflag_x_sub1_min),
        "class_name_prefix": overrides.get("class_name_prefix", base.class_name_prefix),
        "category_name_prefix": overrides.get("category_name_prefix", base.category_name_prefix),
    }
    return LangConfig(**data)
