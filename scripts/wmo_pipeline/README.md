# WMO Table Pipeline

Modular, config-driven framework for extracting WMO code tables from
multilingual PDFs.  Supports **BUFR**, **GRIB**, and **CCT** standards
across **English, French, Spanish, and Russian**.

One config file.  One command.  All tables extracted, validated, aligned,
and enriched with notes.

```
                  YAML config
                      │
             ┌────────▼────────┐
             │   run_pipeline  │
             └────┬───┬───┬───┬┘
                  │   │   │   │
            Extract Validate Align Notes
                  │   │   │   │
            ┌─────▼───▼───▼───▼─────┐
            │     Plugin Registry    │
            │  BUFR · GRIB · CCT    │
            └───────────────────────┘
```

---

## Prerequisites

```bash
pip install -r requirements.txt
```

Python 3.10+ required.

### Download source PDFs

The WMO-306 PDFs are not included in this repo (111 MB total).
Download them from the [WMO website](https://library.wmo.int/) and
place them in `data/`:

```
data/
├── en/   306_I2_2019_updated_2021_en.pdf
├── es/   306-I-2_2019_updated_2021_es.pdf
├── fr/   306_I2_2019_fr.pdf
└── ru/   306_I2_2019_ru.pdf
```

### English reference CSVs

For the alignment and notes steps, place English reference CSVs in
`data/en_reference/`.  These are the BUFR tables from the latest WMO
edition in structured CSV format.

---

The framework is **fully self-contained** — the BUFR extraction engine
is built into `wmo/_bufr_engine/`.  No external dependencies on other
project directories.

---

## Quickstart

### 1. Pick a config

Pre-built configs live in `configs/`:

| Config | Standard | Language | Edition |
|--------|----------|----------|---------|
| `bufr_fr_2019.yaml` | BUFR | French | 2019 |
| `bufr_es_2021.yaml` | BUFR | Spanish | 2019+2021 |
| `bufr_ru_2019.yaml` | BUFR | Russian | 2019 |

### 2. Run

```bash
# Full pipeline — one command:
python run.py --config configs/bufr_fr_2019.yaml

# Dry run first (validate config, show plan):
python run.py --config configs/bufr_fr_2019.yaml --dry-run
```

### 3. Check output

```
output_dir/{lang}/
├── raw/                  ← Step 1: extracted CSVs
├── fixed/                ← Step 2: ID-validated CSVs
├── id_report.csv         ← Step 2: change log
├── aligned/              ← Step 3: alignment CSVs
├── alignment_report.md   ← Step 3: coverage summary
└── final/                ← Step 4: CSVs with notes populated
```

---

## CLI Reference

```bash
# Full pipeline:
python run.py --config configs/bufr_fr_2019.yaml

# Override output directory:
python run.py -c configs/bufr_fr_2019.yaml -o /tmp/test/

# Run only specific steps:
python run.py -c configs/bufr_fr_2019.yaml --steps extract,validate

# Validate config without running:
python run.py -c configs/bufr_fr_2019.yaml --dry-run

# List all registered plugins:
python run.py --list-plugins
```

| Flag | Description |
|------|-------------|
| `--config`, `-c` | Path to YAML/JSON config file (required) |
| `--output-dir`, `-o` | Override output directory from config |
| `--steps` | Override steps: comma-separated `extract,validate,align,notes` |
| `--dry-run` | Show plan without running |
| `--list-plugins` | Show registered extractors, validators, aligners, notes |

---

## Config File Format

Each job is described by a single YAML file.  Copy `configs/_template.yaml`
and fill in your values.  All paths can be absolute or relative to the
config file's directory.

```yaml
# ── Required ──────────────────────────────────────────────
name: "BUFR French 2019"
standard: bufr                    # bufr | grib | cct
lang: fr                          # en | fr | es | ru
output_dir: "../output/fr/"

# ── Source (one of pdf_path or input_dir) ─────────────────
pdf_path: "../data/fr/306_I2_2019_fr.pdf"
page_ranges:
  a: [252, 253]                   # 0-indexed [start, end]
  b: [256, 405]
  c: [406, 411]
  d: [412, 713]
  codeflag: [714, 936]

# OR: use pre-extracted CSVs
# input_dir: "/path/to/existing/csvs/"

# ── Reference (for align and notes steps) ─────────────────
ref_dir: "../data/en_reference/"

# ── Pipeline control ─────────────────────────────────────
steps:                             # default: all four
  - extract
  - validate
  - align
  - notes

tables:                            # default: all for the standard
  - all

# ── Optional ──────────────────────────────────────────────
edition: "2019"
# copy_to: "/path/to/destination/"
# lang_overrides:                  # override LangConfig fields
#   col_bounds_b:
#     fxy: [85, 150]
```

### Page range keys

Short keys are aliases for canonical names:

| Short | Canonical |
|-------|-----------|
| `a` | `table_a` |
| `b` | `table_b` |
| `c` | `table_c` |
| `d` | `table_d` |
| `codeflag` | `codeflag` |

### Nested YAML style

The config loader also accepts a nested style — use whichever
you prefer:

```yaml
job:
  name: "BUFR French 2019"
  standard: bufr
  lang: fr
source:
  pdf: "/path/to/pdf"
  page_ranges: {a: [252, 253], b: [256, 405]}
reference:
  dir: "/path/to/ref/"
output:
  dir: "/path/to/output/"
  copy_to: "/path/to/final/destination/"
pipeline:
  steps: [extract, validate, align, notes]
  tables: [all]
```

---

## Pipeline Steps

### Step 1: Extract (PDF → raw CSVs)

Reads the source PDF with pymupdf and produces one CSV per class,
category, or table.  Position-based extraction uses span bounding boxes
(`x0` for column, `y0` for row) — far more reliable than text-based
parsing for complex multi-column WMO tables.

### Step 2: Validate (raw → fixed CSVs)

Checks every row's `Id` column against the WMO ID format spec.
Invalid IDs are reconstructed from FXY/CodeFigure columns.
Produces `id_report.csv` listing all changes.

BUFR ID formats:

| Table | Format | Example |
|-------|--------|---------|
| A | `bufr4/a/{CodeFigure}` | `bufr4/a/0` |
| B | `bufr4/{F}/{XX}/{YYY}` | `bufr4/0/01/003` |
| C | `bufr4/{FXY}` | `bufr4/201YYY` |
| D | `bufr4/3/{XX}/{YYY}/{occ}/{FXY2}` | `bufr4/3/01/001/1/007001` |
| CodeFlag | `bufr4/{F}/{XX}/{YYY}/{cf}` | `bufr4/0/02/002/All-4` |

### Step 3: Align (fixed vs reference → report)

Compares each language CSV against the English reference on the `Id`
column.  Categorizes rows as `MATCH`, `MISSING_IN_LANG`, `EXTRA_IN_LANG`,
or `ALL_N_MISMATCH`.  Writes per-file alignment CSVs and a markdown
summary report with coverage percentages.

### Step 4: Notes (fixed + reference → final CSVs)

Matches language rows by `Id` to English reference, copying `noteIDs`
and translating `Note_en` → `Note_{lang}` using language-specific rules
(e.g. "see" → "voir", "and" → "et" for French).

---

## Plugin Architecture

The framework uses a **registry pattern** to map `(standard, table_type)`
to handler classes.  Each plugin self-registers at import time.

### Registered Plugins

```
$ python run.py --list-plugins

BUFR: tables=[table_a, table_b, table_c, table_d, codeflag],
      validator=yes, aligner=yes, notes=yes
CCT:  tables=[generic],
      validator=no, aligner=yes, notes=yes
GRIB: tables=[parameter, level, template],
      validator=no, aligner=yes, notes=yes
```

BUFR is fully implemented.  GRIB and CCT are skeleton plugins
(raise `NotImplementedError` if invoked).

### Base Classes

| ABC | Module | Methods |
|-----|--------|---------|
| `BaseExtractor` | `wmo/extractors/__init__.py` | `extract()`, `save()` |
| `BaseValidator` | `wmo/validators/__init__.py` | `validate_file()`, `validate_directory()` |
| `BaseAligner` | `wmo/aligners/__init__.py` | `align_directory()`, `write_report()` |
| `BaseNotesProcessor` | `wmo/notes/__init__.py` | `populate()` |

### How BUFR works

BUFR extractors are thin adapters (~10 lines each) that delegate to the
proven extraction functions (built into `wmo/_bufr_engine/`).  This means:

- **Zero regression risk** — 1500+ lines of battle-tested code are untouched
- The same extraction logic that produced the verified baselines
  (80 files, 14,039 rows for French) is used directly
- The adapter pattern lets the framework focus on orchestration while
  the extraction engine focuses on PDF parsing

---

## How to Add a New Standard (GRIB, CCT)

### 1. Create extractor(s)

```python
# wmo/extractors/grib.py

from wmo.extractors import BaseExtractor
from wmo.registry import registry

class GribParameterExtractor(BaseExtractor):
    standard = "grib"
    table_type = "parameter"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        # Your pymupdf extraction logic here
        return {"00": df_section_00, "01": df_section_01}

    def save(self, results, output_dir, lang):
        total = 0
        for key, df in results.items():
            path = Path(output_dir) / f"GRIB_Parameter_{lang}_{key}.csv"
            df.to_csv(path, index=False)
            total += len(df)
        return total

registry.register_extractor("grib", "parameter", GribParameterExtractor)
```

### 2. (Optional) Create validator

Only needed if the new standard has a different ID format than BUFR:

```python
# wmo/validators/grib.py

from wmo.validators import BaseValidator
from wmo.registry import registry

class GribIdValidator(BaseValidator):
    def validate_file(self, path, output_dir=None):
        ...
    def validate_directory(self, input_dir, output_dir):
        ...

registry.register_validator("grib", GribIdValidator)
```

### 3. Register in auto_discover

Add the import to `wmo/registry.py`:

```python
def auto_discover():
    ...
    _try_import("wmo.extractors.grib")
    _try_import("wmo.validators.grib")   # if created
```

### 4. Create config

```yaml
# configs/grib_fr_2019.yaml
name: "GRIB French 2019"
standard: grib
lang: fr
pdf_path: "/path/to/grib_pdf.pdf"
page_ranges:
  parameter: [10, 50]
  level: [51, 70]
output_dir: "/path/to/output/"
steps: [extract, validate]
tables: [all]
```

### 5. Run

```bash
python run.py --config configs/grib_fr_2019.yaml
```

### What you get for free

These components are **standard-agnostic** and work for any new standard
without modification:

- Config loading and validation
- Path resolution (relative paths)
- Pipeline orchestration and progress reporting
- Timing and summary output
- Reference alignment (if CSVs use an `Id` column)
- Notes population (if CSVs use `noteIDs` / `Note_{lang}` columns)
- `--dry-run`, `--list-plugins`, `--steps` overrides
- `copy_to` for result distribution

---

## How to Add a New Language

### 1. Calibrate column boundaries

Open a known page in the new language's PDF and inspect span positions:

```python
import fitz
doc = fitz.open("path/to/new_lang.pdf")
page = doc[256]  # a Table B page
for block in page.get_text("dict")["blocks"]:
    for line in block.get("lines", []):
        for span in line["spans"]:
            print(f"x0={span['bbox'][0]:.0f}  text={span['text'][:40]}")
```

Use the `x0` values to define column boundaries:

```python
_COL_BOUNDS_B_ZH = {
    "fxy": (85, 150),
    "name": (150, 310),
    "bufr_unit": (310, 400),
    # ... measure for each column
}
```

### 2. Add LangConfig

In `wmo/_bufr_engine/lang_config.py`:

```python
LANG_CONFIGS["zh"] = LangConfig(
    code="zh",
    table_a_start_markers=["..."],
    table_a_end_markers=["..."],
    class_pattern=_compile(r"类\s+(\d+)\s*[—\-–]\s*(.+)"),
    category_pattern=_compile(r"..."),
    note_pattern_b=_compile(r"..."),
    note_pattern_c=_compile(r"..."),
    col_bounds_b=_COL_BOUNDS_B_ZH,
    col_bounds_c=_COL_BOUNDS_C_ZH,
    col_bounds_d=_COL_BOUNDS_D_ZH,
    tab_separated_b=True,      # inspect whether new PDF merges spans
    continuation_filter="续",   # continuation substring in headings
    codeflag_footer_pattern=_compile(r"..."),
    codeflag_continuation_pattern=_compile(r"..."),
    codeflag_marker_phrases=frozenset({"..."}),
    codeflag_footer_texts=frozenset({"FM 94 BUFR", "..."}),
    codeflag_footer_keyword="...",
    codeflag_x_fxy_center=250.0,
    codeflag_x_code_fig_max=110.0,
    codeflag_x_entry_name_min=150.0,
    codeflag_x_sub1_min=390.0,
)
```

### 3. Add note translator

In `wmo/_bufr_engine/populate_notes.py`:

```python
def _translate_note_en_to_zh(note: str) -> str:
    if not note or (isinstance(note, float) and pd.isna(note)):
        return ""
    note = str(note).strip()
    note = re.sub(r"\bsee\b", "见", note, flags=re.IGNORECASE)
    note = re.sub(r"\band\b", "和", note)
    return note

_TRANSLATORS["zh"] = _translate_note_en_to_zh
```

### 4. Create config

```yaml
# configs/bufr_zh_2019.yaml
name: "BUFR Chinese 2019"
standard: bufr
lang: zh
pdf_path: "/path/to/306_I2_2019_zh.pdf"
page_ranges:
  a: [250, 252]
  b: [255, 400]
  # ... calibrate page numbers
ref_dir: "/path/to/english_reference/"
output_dir: "/path/to/output/"
steps: [extract, validate, align, notes]
tables: [all]
```

### 5. Test incrementally

```bash
# Extract just Table A first:
python run.py -c configs/bufr_zh_2019.yaml --steps extract,validate

# Then add more tables once Table A works.
```

---

## History

This pipeline evolved from a BUFR-only extraction tool.  The proven BUFR
extraction engine (~1500 lines) is incorporated directly in `wmo/_bufr_engine/`,
making the framework fully self-contained.  The plugin architecture was added
to support GRIB, CCT, and future WMO standards.

---

## File Structure

```
wmo_pipeline/
├── run.py                          # CLI entry point
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── TUTORIAL.md                     # Step-by-step guide
├── REPORT.md                       # Architecture deep-dive
├── configs/
│   ├── _template.yaml              # Annotated template (copy this)
│   ├── bufr_fr_2019.yaml           # French BUFR
│   ├── bufr_es_2021.yaml           # Spanish BUFR
│   └── bufr_ru_2019.yaml           # Russian BUFR
├── data/                           # PDF source files (not in git)
│   ├── en/                         # English reference PDF
│   ├── es/                         # Spanish PDF
│   ├── fr/                         # French PDF
│   └── ru/                         # Russian PDF
├── tools/
│   ├── docling_check.py            # Cross-check with Docling AI
│   ├── FINDINGS.md                 # Docling evaluation results
│   └── README.md                   # Standalone tool docs
└── wmo/
    ├── __init__.py                 # Package init
    ├── config.py                   # JobConfig + load_config()
    ├── registry.py                 # Plugin registry + auto_discover()
    ├── runner.py                   # Pipeline orchestrator
    ├── _bufr_engine/               # Built-in BUFR extraction engine
    │   ├── lang_config.py          # LangConfig for EN/FR/ES/RU
    │   ├── id_utils.py             # ID validation + reconstruction
    │   ├── populate_notes.py       # Notes population
    │   ├── extract/                # PDF → CSV extractors
    │   │   ├── extract_abc.py      # Tables A, B, C, CodeFlag
    │   │   └── extract_d.py        # Table D
    │   └── align/                  # Reference alignment
    │       └── align_to_reference.py
    ├── extractors/
    │   ├── __init__.py             # BaseExtractor ABC
    │   ├── bufr.py                 # 5 BUFR extractors (adapters)
    │   ├── grib.py                 # 3 GRIB skeletons
    │   └── cct.py                  # 1 CCT skeleton
    ├── validators/
    │   ├── __init__.py             # BaseValidator ABC
    │   └── bufr.py                 # BUFR ID validator
    ├── aligners/
    │   ├── __init__.py             # BaseAligner ABC
    │   └── reference.py            # Reference aligner (all standards)
    └── notes/
        ├── __init__.py             # BaseNotesProcessor ABC
        └── reference.py            # Notes processor (all standards)
```

---

## Known Baselines

| Language | Edition | Files | Rows | ID Fixes | Unresolved |
|----------|---------|-------|------|----------|------------|
| French | 2019 | 80 | 14,039 | 115 | 0 |
| Spanish | 2019+2021 | 79 | ~14,000 | 92 | 0 |
| Russian | 2019 | 78 | 14,424 | 113 | 0 |

Coverage gaps vs English 2025 reference reflect PDF edition differences,
not extraction bugs.

---

## Troubleshooting

### `ImportError: PyYAML is required`

```bash
pip install pyyaml
```

### Config validation errors

Run with `--dry-run` to see the full resolved config and any errors:

```bash
python run.py -c configs/bufr_fr_2019.yaml --dry-run
```

### `lang_overrides` for different PDF layout

If your PDF has different column positions, override them in the config
instead of modifying `lang_config.py`:

```yaml
lang_overrides:
  col_bounds_b:
    fxy: [60, 135]
    name: [135, 350]
    bufr_unit: [350, 440]
    bufr_scale: [440, 490]
    bufr_ref: [490, 560]
    bufr_width: [560, 630]
    crex_unit: [630, 700]
    crex_scale: [700, 745]
    crex_width: [745, 790]
```

---

## Dependencies

| Package | Purpose | Required for |
|---------|---------|--------------|
| `pymupdf` (fitz) | PDF text extraction | All extraction |
| `pymupdf4llm` | PDF → markdown | Table A extraction |
| `pandas` | CSV manipulation | All steps |
| `pyyaml` | YAML config parsing | Config loading |

---

## Agent Documentation

For AI agents (Claude, etc.) working on this codebase:

- **[CLAUDE.md](CLAUDE.md)** — Project rules, environment setup, and key constraints (auto-loaded by Claude Code)
- **[docs/AGENT_GUIDE.md](docs/AGENT_GUIDE.md)** — Deep technical reference: data flow, plugin contracts, CSV schemas, engine internals, known gotchas
