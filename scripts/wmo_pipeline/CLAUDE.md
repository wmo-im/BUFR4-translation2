# CLAUDE.md — Agent Instructions for wmo_pipeline

## What This Project Is

WMO Table Pipeline: extracts structured CSV tables from multilingual WMO-306 PDF documents.
Supports BUFR (fully implemented), GRIB and CCT (skeleton plugins).
Languages: English, French, Spanish, Russian.

## Environment Setup

```bash
# Activate the correct Python environment before running anything:
source ~/.zshrc 2>/dev/null; work wmo;

# Install dependencies (always use uv, never bare pip):
uv pip install -r requirements.txt
```

Python 3.10+ required. Dependencies: pymupdf, pymupdf4llm, pandas, pyyaml.

## How to Run

```bash
# Full pipeline (one command):
python run.py --config configs/bufr_fr_2019.yaml

# Dry run (validate config only):
python run.py --config configs/bufr_fr_2019.yaml --dry-run

# Specific steps only:
python run.py --config configs/bufr_fr_2019.yaml --steps extract,validate

# List registered plugins:
python run.py --list-plugins
```

## Project Structure (Key Files)

```
wmo_pipeline/
├── run.py                              # CLI entry point — start here
├── configs/
│   ├── _template.yaml                  # Copy this for new jobs
│   ├── bufr_fr_2019.yaml              # French BUFR config
│   ├── bufr_es_2021.yaml              # Spanish BUFR config
│   └── bufr_ru_2019.yaml              # Russian BUFR config
├── wmo/
│   ├── config.py                       # load_config() → JobConfig dataclass
│   ├── registry.py                     # Plugin registry + auto_discover()
│   ├── runner.py                       # Pipeline orchestrator (4 steps)
│   ├── extractors/
│   │   ├── __init__.py                 # BaseExtractor ABC
│   │   └── bufr.py                     # 5 BUFR extractors (thin adapters)
│   ├── validators/
│   │   ├── __init__.py                 # BaseValidator ABC
│   │   └── bufr.py                     # BUFR ID validator
│   ├── aligners/
│   │   ├── __init__.py                 # BaseAligner ABC
│   │   └── reference.py               # Reference aligner (all standards)
│   ├── notes/
│   │   ├── __init__.py                 # BaseNotesProcessor ABC
│   │   └── reference.py               # Notes processor (all standards)
│   └── _bufr_engine/                   # Core extraction engine (~1500 lines)
│       ├── lang_config.py              # LangConfig per language (column bounds, patterns)
│       ├── id_utils.py                 # ID validation + reconstruction
│       ├── populate_notes.py           # Notes population from English reference
│       ├── extract/
│       │   ├── extract_abc.py          # Tables A, B, C, CodeFlag extraction
│       │   └── extract_d.py            # Table D extraction
│       └── align/
│           └── align_to_reference.py   # Alignment vs English reference
├── tools/
│   └── docling_check.py               # Cross-check extraction with Docling AI
├── align_cct.py                        # Standalone CCT alignment (1160 lines)
└── fix_cct_issues.py                   # CCT-specific post-processing fixes
```

## Pipeline Steps (in order)

1. **Extract** — PDF → raw CSVs (position-based pymupdf parsing)
2. **Validate** — Fix malformed IDs, produce `id_report.csv`
3. **Align** — Compare language CSVs vs English reference, produce `alignment_report.md`
4. **Notes** — Copy `noteIDs` from English reference into language CSVs

Output structure: `{output_dir}/{lang}/raw/`, `fixed/`, `aligned/`, `final/`

## Architecture Decisions

- **Position-based extraction**: Uses pymupdf `page.get_text("dict")` span bounding boxes (x for column, y for row). Far more reliable than text-based parsing for multi-column WMO tables.
- **Adapter pattern**: BUFR extractors are ~10-line adapters over the proven `_bufr_engine/`. Do NOT refactor or inline the engine — it is battle-tested.
- **Self-registering plugins**: Each plugin module registers at import time. `auto_discover()` in registry.py triggers all imports.
- **Config-driven**: All user-facing parameters are in YAML configs. No Python editing for standard workflows.

## Critical Rules

- NEVER modify `wmo/_bufr_engine/` internals without understanding the full extraction pipeline. The engine has language-specific calibrations (column boundaries, regex patterns) that are tightly coupled to specific PDF layouts.
- Column boundaries in `lang_config.py` are calibrated per-PDF by measuring span x-coordinates. Changing them without re-measuring will break extraction.
- ID formats are strict (see `id_utils.py`). The `bufr4/...` prefix is mandatory.
- Coverage gaps between language editions and the English 2025 reference are expected — they reflect PDF edition differences, not bugs.
- Table D uses a per-(FXY1, FXY2) occurrence counter (`seq`), not a simple positional counter. This is a known divergence from English that gets fixed in post-processing.
- CodeFlag tables use centered FXY headers and text-pattern matching, not fixed x-boundaries.
- Always use `uv pip install`, never bare `pip install`.

## YAML Config Schema (quick reference)

Required: `name`, `standard` (bufr/grib/cct), `lang` (en/fr/es/ru), `output_dir`
Source: `pdf_path` + `page_ranges` OR `input_dir` (pre-extracted CSVs)
Optional: `ref_dir`, `steps`, `tables`, `edition`, `copy_to`, `lang_overrides`
Page ranges use 0-indexed pages. Short keys: a→table_a, b→table_b, c→table_c, d→table_d.

## Testing Changes

After any code change:
1. Run dry-run: `python run.py -c configs/bufr_fr_2019.yaml --dry-run`
2. Run full pipeline: `python run.py -c configs/bufr_fr_2019.yaml`
3. Verify output counts match known baselines:
   - French: 80 files, 14,039 rows, 115 ID fixes
   - Spanish: 79 files, ~14,000 rows, 92 ID fixes
   - Russian: 78 files, 14,424 rows, 113 ID fixes

## Common Tasks

- **Add new language**: Add LangConfig in `lang_config.py`, note translator in `populate_notes.py`, YAML config in `configs/`
- **Add new standard**: Create extractor in `wmo/extractors/`, register in `registry.py`, create YAML config
- **Override column positions**: Use `lang_overrides` in YAML, not code changes
- **Cross-check extraction**: Use `tools/docling_check.py` for AI-based comparison
