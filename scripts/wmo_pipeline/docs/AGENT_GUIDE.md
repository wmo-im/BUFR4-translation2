# Agent Technical Reference — WMO Table Pipeline

> This document is for AI agents (Claude, etc.) working on this codebase.
> It provides the technical context needed to understand, modify, and extend the pipeline
> without reading every source file. For human-oriented docs, see README.md and TUTORIAL.md.

---

## 1. System Overview

**Purpose:** Extract structured CSV tables from WMO-306 multilingual PDFs using position-based parsing.

**Execution model:** YAML config → `run.py` CLI → 4-step pipeline (extract → validate → align → notes) → CSV output.

**Standards supported:** BUFR (fully implemented), GRIB (skeleton), CCT (skeleton).

**Languages:** en, fr, es, ru. Each language has a calibrated `LangConfig` with column boundaries measured from the specific PDF.

---

## 2. Entry Points

| What | Where | How |
|------|-------|-----|
| CLI | `run.py` | `python run.py --config configs/bufr_fr_2019.yaml` |
| Config loader | `wmo/config.py:load_config()` | Returns `JobConfig` dataclass |
| Pipeline runner | `wmo/runner.py:run_pipeline()` | Orchestrates all 4 steps |
| Plugin discovery | `wmo/registry.py:auto_discover()` | Imports all plugin modules |
| BUFR engine init | `wmo/__init__.py:ensure_bufr_engine()` | Adds `_bufr_engine/` to sys.path |

---

## 3. Data Flow (Detailed)

```
PDF file (WMO-306, ~900 pages, multilingual)
    │
    ▼ pymupdf page.get_text("dict") → spans with (x, y, text)
    │
    ▼ Column classification by x-coordinate (LangConfig.col_bounds_*)
    │ Row grouping by y-coordinate (±3pt = same row)
    │
    ▼ Step 1: EXTRACT → raw/*.csv
    │   Each table type produces 1+ CSV files
    │   Table B: one CSV per class (00-40), e.g. BUFRCREX_TableB_fr_01.csv
    │   Table D: one CSV per category, e.g. BUFR_TableD_fr_01.csv
    │   CodeFlag: one CSV per class
    │   Table A: single CSV
    │   Table C: single CSV
    │
    ▼ Step 2: VALIDATE → fixed/*.csv + id_report.csv
    │   Checks Id column against format spec
    │   Reconstructs invalid IDs from FXY/CodeFigure columns
    │   Logs all changes to id_report.csv
    │
    ▼ Step 3: ALIGN → aligned/*.csv + alignment_report.md
    │   Joins language CSV with English reference on Id
    │   Categories: MATCH, MISSING_IN_LANG, EXTRA_IN_LANG, ALL_N_MISMATCH
    │   Alignment CSVs are for analysis only, not final output
    │
    ▼ Step 4: NOTES → final/*.csv
        Copies noteIDs column from English reference (by Id match)
        Translates Note_en → Note_{lang} using basic word substitution
        Final CSVs = fixed CSVs + noteIDs column
```

### CSV Column Schema

**Table B** (`BUFRCREX_TableB_{lang}_{class}.csv`):
```
Id, FXY, EntryName_{lang}, BUFR_Unit, BUFR_Scale, BUFR_ReferenceValue,
BUFR_DataWidth_Bits, CREX_Unit, CREX_Scale, CREX_DataWidth_Characters,
Note_{lang}, noteIDs
```

**Table D** (`BUFR_TableD_{lang}_{category}.csv`):
```
Id, FXY1, FXY2, EntryName_{lang}, Note_{lang}, noteIDs
```

**CodeFlag** (`BUFRCREX_CodeFlag_{lang}_{class}.csv`):
```
Id, FXY, CodeFigure, EntryName_{lang}, Note_{lang}, noteIDs
```

**Table A** (`BUFR_TableA_{lang}.csv`):
```
Id, CodeFigure, EntryName_{lang}, Note_{lang}, noteIDs
```

**Table C** (`BUFR_TableC_{lang}.csv` or `BUFRCREX_TableC_{lang}.csv`):
```
Id, FXY, EntryName_{lang}, Note_{lang}, noteIDs
```

---

## 4. Plugin Registry

### How It Works

```python
# Each plugin self-registers at module import time:
# wmo/extractors/bufr.py (bottom of file):
registry.register_extractor("bufr", "table_a", BufrTableAExtractor)
registry.register_extractor("bufr", "table_b", BufrTableBExtractor)
# ... etc

# auto_discover() in registry.py triggers all imports:
def auto_discover():
    _try_import("wmo.extractors.bufr")
    _try_import("wmo.validators.bufr")
    _try_import("wmo.aligners.reference")
    _try_import("wmo.notes.reference")
    _try_import("wmo.extractors.grib")
    _try_import("wmo.extractors.cct")
```

### Four Registries

| Registry | Key | Value | Lookup |
|----------|-----|-------|--------|
| Extractors | `(standard, table_type)` | ExtractorClass | `registry.get_extractor("bufr", "table_b")` |
| Validators | `standard` | ValidatorClass | `registry.get_validator("bufr")` |
| Aligners | `standard` | AlignerClass | `registry.get_aligner("bufr")` |
| Notes | `standard` | NotesProcessorClass | `registry.get_notes("bufr")` |

### Base Class Contracts

**BaseExtractor** (`wmo/extractors/__init__.py`):
```python
class BaseExtractor(ABC):
    standard: str    # "bufr", "grib", "cct"
    table_type: str  # "table_a", "table_b", etc.

    @abstractmethod
    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        """Returns dict or DataFrame of extracted data."""

    @abstractmethod
    def save(self, results, output_dir, lang):
        """Saves results to CSV(s). Returns total row count (int)."""
```

**BaseValidator** (`wmo/validators/__init__.py`):
```python
class BaseValidator(ABC):
    @abstractmethod
    def validate_file(self, path, output_dir=None):
        """Validates one CSV. Returns (changes_count, rows_count)."""

    @abstractmethod
    def validate_directory(self, input_dir, output_dir):
        """Validates all CSVs in dir. Returns (total_changes, file_count)."""
```

**BaseAligner** (`wmo/aligners/__init__.py`):
```python
class BaseAligner(ABC):
    @abstractmethod
    def align_directory(self, lang_dir, ref_dir, output_dir, lang):
        """Aligns all CSVs vs reference. Returns summary dict."""

    @abstractmethod
    def write_report(self, summary, output_path):
        """Writes alignment report."""
```

**BaseNotesProcessor** (`wmo/notes/__init__.py`):
```python
class BaseNotesProcessor(ABC):
    @abstractmethod
    def populate(self, input_dir, ref_dir, output_dir, lang):
        """Populates noteIDs from reference. Returns file count."""
```

---

## 5. BUFR Engine Internals

### Position-Based Extraction Strategy

The engine reads PDF pages via `fitz.open(pdf_path)` and processes each page:

```python
page = doc[page_num]
blocks = page.get_text("dict")["blocks"]
for block in blocks:
    for line in block.get("lines", []):
        for span in line["spans"]:
            x = span["bbox"][0]   # left edge x-coordinate
            y = span["bbox"][1]   # top edge y-coordinate
            text = span["text"]
            # Classify into column by x-position using LangConfig bounds
```

### LangConfig (wmo/_bufr_engine/lang_config.py)

Each language has a `LangConfig` dataclass with:

| Field | Purpose | Example (French) |
|-------|---------|-------------------|
| `code` | Language code | `"fr"` |
| `col_bounds_b` | Table B column x-ranges | `{"fxy": (85, 150), "name": (150, 310), ...}` |
| `col_bounds_c` | Table C column x-ranges | `{...}` |
| `col_bounds_d` | Table D column x-ranges | `{...}` |
| `class_pattern` | Regex for class headers | `r"Classe\s+(\d+)"` |
| `category_pattern` | Regex for category headers | `r"Catégorie\s+(\d+)"` |
| `note_pattern_b` | Regex for note references | `r"\(voir Note (\d+)\)"` |
| `continuation_filter` | Text in page-continuation headers | `"suite"` |
| `codeflag_*` | Various CodeFlag-specific parameters | (centered FXY x-pos, code figure max x, etc.) |
| `tab_separated_b` | Whether Table B FXY is tab-separated | `True` for French |

### ID Validation (wmo/_bufr_engine/id_utils.py)

Key functions:
- `is_valid_id(id_str)` — checks `bufr4/...` format
- `detect_table_type(filename)` — infers type from CSV filename stem
- `reconstruct_id(row, table_type)` — builds ID from FXY/CodeFigure columns
- `fix_file(csv_path, output_dir)` — validates and fixes all IDs in one file
- `fix_directory(input_dir, output_dir)` — processes all CSVs

Language-specific normalizations:
- Spanish: "Todo N" → "All-N", "Los N" → "All-N"
- French: "Tous N" / "Toutes N" / "Toute N" → "All-N", "4 bits mis à 1" → "All-4"
- Russian: "Все N" → "All-N"

### Extraction Functions (wmo/_bufr_engine/extract/)

**extract_abc.py:**
- `extract_table_a(pdf, start, end, lang, cfg)` — uses pymupdf4llm markdown (simpler table)
- `extract_table_b(pdf, start, end, lang, cfg)` — position-based, multi-file by class
- `extract_table_c(pdf, start, end, lang, cfg)` — position-based, single file
- `extract_codeflag(pdf, start, end, lang, cfg)` — position-based, multi-file by class

**extract_d.py:**
- `extract_table_d(pdf, start, end, lang, cfg)` — position-based, multi-file by category

### Alignment (wmo/_bufr_engine/align/align_to_reference.py)

- `align_directory(lang_dir, ref_dir, output_dir, lang)` — outer join on Id column
- `_match_filename(lang_file, ref_files)` — maps `*_{lang}_*.csv` → `*_en_*.csv`
  - Handles both `_{lang}_` (multi-file) and `_{lang}.` (single-file) suffixes
  - Handles `BUFRCREX_` ↔ `BUFR_` prefix swap for Table C

---

## 6. Configuration Deep Dive

### JobConfig Fields (wmo/config.py)

```python
@dataclass
class JobConfig:
    name: str               # "BUFR French 2019"
    standard: str           # "bufr" | "grib" | "cct"
    lang: str               # "en" | "fr" | "es" | "ru"
    output_dir: Path        # root output directory
    edition: str = ""       # "2019", "2019+2021"
    pdf_path: Path = None   # source PDF (if extracting)
    input_dir: Path = None  # pre-extracted CSVs (if skipping extract)
    ref_dir: Path = None    # English reference CSVs
    page_ranges: dict = {}  # {table_key: [start, end]} 0-indexed
    steps: list = ["extract", "validate", "align", "notes"]
    tables: list = ["all"]  # or specific table types
    lang_overrides: dict = {}  # override LangConfig fields
    copy_to: Path = None    # copy final results here
```

### Path Resolution

All paths in YAML are resolved relative to the config file's directory.
`_resolve_path(value, config_dir)` handles this automatically.

### Page Range Aliases

Short keys are expanded: `a` → `table_a`, `b` → `table_b`, `c` → `table_c`, `d` → `table_d`.

### Nested YAML

The loader supports nested style via `_flatten_config()`:
```yaml
job:
  name: "..."
  standard: bufr
source:
  pdf: "..."
  page_ranges: {a: [252, 253]}
```
→ flattened to top-level keys before creating JobConfig.

---

## 7. Common Modification Patterns

### Adding a new extractor for an existing standard

1. Create class inheriting `BaseExtractor` with `standard` and `table_type` attrs
2. Implement `extract()` and `save()` methods
3. Call `registry.register_extractor(standard, table_type, YourClass)` at module level
4. Add `_try_import("wmo.extractors.your_module")` to `auto_discover()` in `registry.py`
5. Add the table type to your YAML config's `page_ranges`

### Adding a new language

1. Measure column x-coordinates from the PDF (see TUTORIAL.md for method)
2. Create `_COL_BOUNDS_*_{LANG}` dicts in `lang_config.py`
3. Add `LangConfig` to `LANG_CONFIGS` dict with all patterns calibrated
4. Add note translator function to `_TRANSLATORS` dict in `populate_notes.py`
5. Create YAML config in `configs/`

### Changing column boundaries without code changes

Use `lang_overrides` in the YAML config:
```yaml
lang_overrides:
  col_bounds_b:
    fxy: [60, 135]
    name: [135, 350]
```

### Skipping steps

Set `steps` in YAML or use `--steps` CLI flag:
```bash
python run.py -c config.yaml --steps validate,align
```
Requires that previous step output already exists (e.g., `raw/` must exist to run validate).

---

## 8. Known Gotchas

### FXY span formats
FXY codes in PDFs appear as either "F XX YYY" in one span OR split as "F XX" + "YYY" in two spans. Both must be handled. French Table B has 3 variants: tab-separated, tab FXY-only, space-separated — handled by `_split_merged_spans()` in `extract_abc.py`.

### CodeFlag detection
CodeFlag tables have centered FXY headers (x~272), not columnar. Use text pattern matching (`CODE_FIG_PATTERN`) rather than fixed x-boundaries. Marker detection uses exact phrase matching ("Кодовая", "цифра", "Номер бита") with x-position check, NOT individual words.

### Footer filtering
Use text-based detection (check for "Кодовые таблицы/Таблицы флагов", "I.2 –", etc.) rather than y-position cutoffs. Content extends to y~738 and footers start at y~758.

### Continuation pages
Parentheses around continuation markers may be split into separate spans. Use optional parens in regex: `\(?pattern\)?`.

### Table D counter scheme
Pipeline `extract_d.py` assigns `seq` via per-(FXY1, FXY2) `cumcount() + 1`. This does NOT match English — English uses a positional counter. Post-processing is needed to align Table D IDs.

### Edition differences
French PDF is strictly 2019 edition. Spanish is 2019+2021. English reference is 2025 edition. Coverage gaps (e.g., French Table B 91% vs English) are edition differences, not bugs.

### File matching for alignment
`_match_filename()` must handle both `_{lang}_` (multi-file) and `_{lang}.` (single-file) suffix patterns, plus `BUFRCREX_` ↔ `BUFR_` prefix swap for Table C.

---

## 9. Verified Baselines

Use these to validate that code changes haven't introduced regressions:

| Language | Edition | Files | Total Rows | ID Fixes | Unresolved |
|----------|---------|-------|------------|----------|------------|
| French | 2019 | 80 | 14,039 | 115 | 0 |
| Spanish | 2019+2021 | 79 | ~14,000 | 92 | 0 |
| Russian | 2019 | 78 | 14,424 | 113 | 0 |

---

## 10. File-Level Reference

### Core Pipeline

| File | Lines | Key Exports |
|------|-------|-------------|
| `run.py` | ~113 | `main()` — CLI arg parsing, calls `load_config()` + `run_pipeline()` |
| `wmo/config.py` | ~254 | `load_config(path) → JobConfig`, `JobConfig` dataclass |
| `wmo/registry.py` | ~135 | `registry` (singleton), `auto_discover()`, `Registry` class |
| `wmo/runner.py` | ~317 | `run_pipeline(config)`, `_run_extraction()`, `_run_validation()`, `_run_alignment()`, `_run_notes()` |

### Extractors

| File | Lines | Registers |
|------|-------|-----------|
| `wmo/extractors/__init__.py` | ~103 | `BaseExtractor` ABC |
| `wmo/extractors/bufr.py` | ~124 | 5 extractors: table_a, table_b, table_c, table_d, codeflag |
| `wmo/extractors/grib.py` | ~90 | 3 skeletons: parameter, level, template |
| `wmo/extractors/cct.py` | ~65 | 1 skeleton: generic |

### Validators, Aligners, Notes

| File | Lines | Registers |
|------|-------|-----------|
| `wmo/validators/bufr.py` | ~40 | `BufrIdValidator` for "bufr" |
| `wmo/aligners/reference.py` | ~41 | `ReferenceAligner` for bufr, grib, cct |
| `wmo/notes/reference.py` | ~30 | `ReferenceNotesProcessor` for bufr, grib, cct |

### BUFR Engine

| File | Lines | Key Functions |
|------|-------|---------------|
| `wmo/_bufr_engine/lang_config.py` | ~300 | `get_lang_config(lang, overrides)`, `LANG_CONFIGS` dict, `LangConfig` dataclass |
| `wmo/_bufr_engine/id_utils.py` | ~250 | `is_valid_id()`, `reconstruct_id()`, `fix_file()`, `fix_directory()` |
| `wmo/_bufr_engine/populate_notes.py` | ~152 | `populate_notes()`, `_TRANSLATORS` dict |
| `wmo/_bufr_engine/extract/extract_abc.py` | ~500 | `extract_table_a()`, `extract_table_b()`, `extract_table_c()`, `extract_codeflag()` |
| `wmo/_bufr_engine/extract/extract_d.py` | ~412 | `extract_table_d()` |
| `wmo/_bufr_engine/align/align_to_reference.py` | ~310 | `align_directory()`, `write_summary_report()`, `_match_filename()` |

### Standalone Tools

| File | Lines | Purpose |
|------|-------|---------|
| `align_cct.py` | ~1160 | CCT alignment with per-table functions (align_c00 through align_c14) |
| `fix_cct_issues.py` | ~200 | CCT-specific fixes: header contamination, truncation |
| `tools/docling_check.py` | varies | Cross-check extraction against Docling AI model |
