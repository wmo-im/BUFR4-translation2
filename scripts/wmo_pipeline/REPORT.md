# WMO Table Pipeline — Architecture Report

## 1. Executive Summary

The **WMO Table Pipeline** (`wmo_pipeline/`) is a modular, config-driven framework for extracting WMO code tables from multilingual PDFs. It wraps the proven `wmo/_bufr_engine/` extraction logic in a plugin architecture that supports **BUFR**, **GRIB**, and **CCT** standards across **English, French, Spanish, and Russian** languages.

**One-button execution:**
```bash
python run.py --config configs/bufr_fr_2019.yaml
```

**Verified output:** 80 CSV files, 14,039 entries, 0 unresolved IDs — matching the established French baseline exactly.

---

## 2. Architecture Overview

```
┌─────────────────────────────────────────────────────────┐
│                     YAML Config                         │
│  (standard, lang, pdf, pages, steps, tables, ref_dir)   │
└──────────────────────────┬──────────────────────────────┘
                           │
                      load_config()
                           │
                    ┌──────▼──────┐
                    │  JobConfig  │
                    └──────┬──────┘
                           │
                     run_pipeline()
                           │
              ┌────────────┼────────────┬──────────────┐
              ▼            ▼            ▼              ▼
         ┌────────┐  ┌──────────┐  ┌───────┐    ┌───────┐
         │EXTRACT │  │ VALIDATE │  │ ALIGN │    │ NOTES │
         └───┬────┘  └────┬─────┘  └───┬───┘    └───┬───┘
             │            │            │             │
     ┌───────▼────────┐   │            │             │
     │    Registry     │   │            │             │
     │ (standard,type) │   │            │             │
     │   → handler     │   │            │             │
     └────────────────┘   │            │             │
             │            │            │             │
     ┌───────▼────────┐   │            │             │
     │ BaseExtractor  │   │            │             │
     │  ├ BufrTableA  │   │            │             │
     │  ├ BufrTableB  │   │            │             │
     │  ├ BufrTableC  │   │            │             │
     │  ├ BufrTableD  │   │            │             │
     │  ├ BufrCodeFlag│   │            │             │
     │  ├ GribParam*  │   │            │             │
     │  └ CctGeneric* │   │            │             │
     └────────────────┘   │            │             │
                          │            │             │
                   ┌──────▼─────┐ ┌────▼─────┐ ┌────▼──────┐
                   │BaseValidator│ │BaseAligner│ │BaseNotes  │
                   │ BufrIdVal  │ │ Reference │ │ Reference │
                   └────────────┘ └──────────┘ └───────────┘
             * = skeleton (NotImplementedError)
```

---

## 3. File Structure

```
wmo_pipeline/
├── run.py                          # CLI entry point
├── REPORT.md                       # This document
├── configs/                        # YAML job configurations
│   ├── _template.yaml              # Annotated template
│   ├── bufr_fr_2019.yaml           # French BUFR (verified ✓)
│   ├── bufr_es_2021.yaml           # Spanish BUFR
│   └── bufr_ru_2019.yaml           # Russian BUFR
└── wmo/                            # Python package
    ├── __init__.py                 # Path setup + version
    ├── config.py                   # JobConfig + load_config()
    ├── registry.py                 # Plugin registry + auto_discover()
    ├── runner.py                   # Pipeline orchestrator (Steps 1–4)
    ├── extractors/
    │   ├── __init__.py             # BaseExtractor ABC
    │   ├── bufr.py                 # 5 BUFR extractors (adapters)
    │   ├── grib.py                 # 3 GRIB extractors (skeleton)
    │   └── cct.py                  # 1 CCT extractor (skeleton)
    ├── validators/
    │   ├── __init__.py             # BaseValidator ABC
    │   └── bufr.py                 # BUFR ID validator (adapter)
    ├── aligners/
    │   ├── __init__.py             # BaseAligner ABC
    │   └── reference.py            # Reference aligner (standard-agnostic)
    └── notes/
        ├── __init__.py             # BaseNotesProcessor ABC
        └── reference.py            # Reference notes processor (standard-agnostic)
```

---

## 4. YAML Configuration Schema

Each pipeline run is driven by a single YAML file. All paths can be absolute or relative to the config file's directory.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `name` | string | yes | Human-readable job name |
| `standard` | string | yes | `bufr`, `grib`, or `cct` |
| `lang` | string | yes | `en`, `fr`, `es`, or `ru` |
| `edition` | string | no | Source document edition (e.g. "2019") |
| `pdf_path` | string | if extracting | Path to source PDF |
| `page_ranges` | dict | if extracting | `{table_key: [start, end]}` (0-indexed) |
| `input_dir` | string | if skipping extract | Pre-extracted CSVs |
| `ref_dir` | string | if aligning/notes | English reference CSV directory |
| `output_dir` | string | yes | Root output directory |
| `steps` | list | no | `[extract, validate, align, notes]` (default: all) |
| `tables` | list | no | `[all]` or specific types (default: all) |
| `lang_overrides` | dict | no | Override LangConfig fields |
| `copy_to` | string | no | Copy final results here |

**Short key aliases** for page_ranges: `a` → `table_a`, `b` → `table_b`, etc.

**Nested YAML style** also supported:
```yaml
job:
  name: "BUFR French 2019"
  standard: bufr
  lang: fr
source:
  pdf: "/path/to/pdf"
  page_ranges: {a: [252, 253]}
reference:
  dir: "/path/to/ref/"
output:
  dir: "/path/to/output/"
pipeline:
  steps: [extract, validate, align, notes]
  tables: [all]
```

---

## 5. Plugin System

### 5.1 Registry

The `Registry` class at `wmo/registry.py` stores four separate registries:
- **Extractors:** `(standard, table_type)` → `ExtractorClass`
- **Validators:** `standard` → `ValidatorClass`
- **Aligners:** `standard` → `AlignerClass`
- **Notes processors:** `standard` → `NotesProcessorClass`

Plugins self-register at import time (module-level code). The runner calls `auto_discover()` at startup to import all known plugin modules.

### 5.2 Current Registrations

```
$ python run.py --list-plugins

Registered plugins:
  BUFR: tables=[table_a, table_b, table_c, table_d, codeflag],
        validator=yes, aligner=yes, notes=yes
  CCT:  tables=[generic],
        validator=no, aligner=yes, notes=yes
  GRIB: tables=[parameter, level, template],
        validator=no, aligner=yes, notes=yes
```

### 5.3 Base Classes

| ABC | Module | Methods to implement |
|-----|--------|---------------------|
| `BaseExtractor` | `wmo/extractors/__init__.py` | `extract(pdf, start, end, lang, cfg)`, `save(results, dir, lang)` |
| `BaseValidator` | `wmo/validators/__init__.py` | `validate_file(path, out)`, `validate_directory(in, out)` |
| `BaseAligner` | `wmo/aligners/__init__.py` | `align_directory(lang, ref, out, lang)`, `write_report(summary, path)` |
| `BaseNotesProcessor` | `wmo/notes/__init__.py` | `populate(input, ref, output, lang)` |

---

## 6. Pipeline Steps (Detail)

### Step 1: Extract (PDF → raw CSVs)

For each table type in the config:
1. Look up `(standard, table_type)` in the extractor registry.
2. Instantiate the extractor class.
3. Call `extractor.extract(pdf_path, start, end, lang, lang_config)`.
4. Call `extractor.save(results, raw_dir, lang)`.

**Output:** `{output_dir}/{lang}/raw/*.csv`

### Step 2: Validate (raw → fixed CSVs)

1. Look up `standard` in the validator registry.
2. Call `validator.validate_directory(raw_dir, fixed_dir)`.
3. Write `id_report.csv` with all changes.

**Output:** `{output_dir}/{lang}/fixed/*.csv` + `id_report.csv`

### Step 3: Align (fixed vs reference → report)

1. Look up `standard` in the aligner registry.
2. Call `aligner.align_directory(fixed_dir, ref_dir, aligned_dir, lang)`.
3. Call `aligner.write_report(summary, report_path)`.

**Output:** `{output_dir}/{lang}/aligned/*.csv` + `alignment_report.md`

### Step 4: Notes (fixed + reference → final CSVs)

1. Look up `standard` in the notes registry.
2. Call `processor.populate(fixed_dir, ref_dir, final_dir, lang)`.

**Output:** `{output_dir}/{lang}/final/*.csv`

### Output Directory Structure

```
{output_dir}/{lang}/
├── raw/                  ← Step 1: extracted CSVs
├── fixed/                ← Step 2: ID-validated CSVs
├── id_report.csv         ← Step 2: change log
├── aligned/              ← Step 3: alignment CSVs (report-only)
├── alignment_report.md   ← Step 3: summary report
└── final/                ← Step 4: CSVs with notes populated
```

---

## 7. BUFR Implementation Details

### 7.1 Adapter Pattern

BUFR extractors are thin adapters (~10 lines each) that delegate to the built-in engine (`wmo/_bufr_engine/`):

```python
class BufrTableBExtractor(BaseExtractor):
    standard = "bufr"
    table_type = "table_b"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        return _extract_table_b(pdf_path, start_page, end_page, lang, lang_config)

    def save(self, results, output_dir, lang):
        return _save_table_b(results, str(output_dir), lang)
```

The `ensure_bufr_engine()` function imports `wmo._bufr_engine` which adds its directory to `sys.path`, enabling the engine's internal bare imports to resolve.

### 7.2 Language Configuration

BUFR extraction uses `LangConfig` (from `wmo/_bufr_engine/lang_config.py`) with per-language:
- **Column x-boundaries** calibrated per PDF
- **Regex patterns** for class/category headings, notes, footers
- **CodeFlag parameters** for centered FXY detection, code figure boundaries

Override via `lang_overrides` in the YAML config.

### 7.3 ID Format

| Table | Format | Example |
|-------|--------|---------|
| A | `bufr4/a/{CodeFigure}` | `bufr4/a/0` |
| B | `bufr4/{F}/{XX}/{YYY}` | `bufr4/0/01/003` |
| C | `bufr4/{FXY}` | `bufr4/201YYY` |
| D | `bufr4/3/{XX}/{YYY}/{occ}/{FXY2}` | `bufr4/3/01/001/1/007001` |
| CodeFlag | `bufr4/{F}/{XX}/{YYY}/{cf}` | `bufr4/0/02/002/All-4` |

### 7.4 Known Baselines

| Language | Edition | Files | Rows | ID Fixes | Coverage vs EN 2025 |
|----------|---------|-------|------|----------|---------------------|
| French | 2019 | 80 | 14,039 | 115 | ~87% (edition gap) |
| Spanish | 2019+2021 | 79 | ~14,000 | 92 | ~92% |
| Russian | 2019 | 78 | 14,424 | 113 | ~88% |

---

## 8. Adding a New Standard (GRIB, CCT)

### Step-by-step guide:

1. **Create extractor file** at `wmo/extractors/grib.py`:
   ```python
   from wmo.extractors import BaseExtractor
   from wmo.registry import registry

   class GribParameterExtractor(BaseExtractor):
       standard = "grib"
       table_type = "parameter"

       def extract(self, pdf_path, start_page, end_page, lang, lang_config):
           # Your extraction logic using pymupdf
           return df

       def save(self, results, output_dir, lang):
           # Save to CSV
           return total_rows

   registry.register_extractor("grib", "parameter", GribParameterExtractor)
   ```

2. **Add to auto_discover()** in `wmo/registry.py`:
   ```python
   _try_import("wmo.extractors.grib")
   ```

3. **Create validator** (if GRIB has different ID format) at `wmo/validators/grib.py`:
   ```python
   from wmo.validators import BaseValidator
   from wmo.registry import registry

   class GribIdValidator(BaseValidator):
       def validate_file(self, path, output_dir=None):
           ...
       def validate_directory(self, input_dir, output_dir):
           ...

   registry.register_validator("grib", GribIdValidator)
   ```

4. **Create YAML config** at `configs/grib_fr_2019.yaml`:
   ```yaml
   name: "GRIB French 2019"
   standard: grib
   lang: fr
   pdf_path: "/path/to/grib_pdf.pdf"
   page_ranges:
     parameter: [10, 50]
   output_dir: "/path/to/output/"
   steps: [extract, validate]
   ```

5. **Run:**
   ```bash
   python run.py --config configs/grib_fr_2019.yaml
   ```

### What's shared (no re-implementation needed):
- **Config loading** — YAML parser, path resolution, validation
- **Registry** — auto-discovery, plugin lookup
- **Runner** — step orchestration, timing, progress reporting
- **Aligner** — reference comparison (if same Id-based CSV structure)
- **Notes processor** — noteIDs/Note_{lang} population (if same schema)
- **Output structure** — raw/fixed/aligned/final directory layout

---

## 9. Adding a New Language

1. **Add LangConfig** in `wmo/_bufr_engine/lang_config.py`:
   ```python
   LANG_CONFIGS["zh"] = LangConfig(
       code="zh",
       class_pattern=_compile(r"类\s+(\d+)\s*[—\-–]\s*(.+)"),
       col_bounds_b=_COL_BOUNDS_B_ZH,  # calibrate per PDF
       ...
   )
   ```

2. **Add note translator** in `wmo/_bufr_engine/populate_notes.py`:
   ```python
   def _translate_note_en_to_zh(note: str) -> str:
       note = re.sub(r"\bsee\b", "见", note, flags=re.IGNORECASE)
       return note

   _TRANSLATORS["zh"] = _translate_note_en_to_zh
   ```

3. **Create YAML config** with new lang and calibrated page ranges.

---

## 10. CLI Reference

```bash
# Full pipeline (one button):
python run.py --config configs/bufr_fr_2019.yaml

# Override output directory:
python run.py --config configs/bufr_fr_2019.yaml --output-dir /tmp/test/

# Override steps (skip extraction, use pre-extracted CSVs):
python run.py --config configs/bufr_fr_2019.yaml --steps validate,align,notes

# Dry run (validate config, show plan):
python run.py --config configs/bufr_fr_2019.yaml --dry-run

# List all registered plugins:
python run.py --list-plugins
```

---

## 11. Dependencies

| Package | Purpose | Required for |
|---------|---------|--------------|
| `pymupdf` (fitz) | PDF text extraction | All extraction |
| `pymupdf4llm` | PDF → markdown (Table A) | Table A extraction |
| `pandas` | CSV manipulation | All steps |
| `pyyaml` | YAML config parsing | Config loading |

---

## 12. Design Decisions

1. **Self-contained engine:** The BUFR extraction engine is built into `wmo/_bufr_engine/`, so the framework has no external directory dependencies. The engine code is the same proven ~1500 lines, just incorporated into the package.

2. **Adapter pattern for BUFR:** Thin adapter classes (~10 lines each) wrap the engine functions behind the `BaseExtractor` interface. This keeps the engine untouched while providing the plugin contract.

3. **Self-registering plugins:** Each plugin module registers itself on import, so the runner doesn't need to know about all implementations.

4. **Standard-agnostic aligner/notes:** The reference aligner and notes processor work for any standard that uses Id-based CSVs — no need to rewrite for GRIB/CCT.

5. **Config-driven, not code-driven:** Everything a user needs to change is in the YAML — no Python editing required for standard workflows.

6. **Relative path resolution:** YAML paths are resolved relative to the config file's directory, making configs portable.
