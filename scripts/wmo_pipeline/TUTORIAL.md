# WMO Table Pipeline — Tutorial

## What this does

You have a WMO-306 PDF in French, Spanish, or Russian.
You want structured CSVs of all BUFR tables (A, B, C, D, CodeFlag).
This pipeline does that in one command.

---

## Step 1: Check your PDF is in `data/`

```
wmo_pipeline/data/
├── en/   306_I2_2019_updated_2021_en.pdf
├── es/   306-I-2_2019_updated_2021_es.pdf
├── fr/   306_I2_2019_fr.pdf
└── ru/   306_I2_2019_ru.pdf
```

If you have a new PDF, drop it in the appropriate `data/{lang}/` folder.

## Step 2: Pick (or create) a config

Pre-built configs are in `configs/`:

```bash
ls configs/
# bufr_fr_2019.yaml    ← French 2019
# bufr_es_2021.yaml    ← Spanish 2019+2021
# bufr_ru_2019.yaml    ← Russian 2019
# _template.yaml       ← copy this for new jobs
```

A config looks like this:

```yaml
name: "BUFR French 2019"
standard: bufr
lang: fr
pdf_path: "../data/fr/306_I2_2019_fr.pdf"
page_ranges:
  a: [252, 253]          # 0-indexed page numbers
  b: [256, 405]
  c: [406, 411]
  d: [412, 713]
  codeflag: [714, 936]
ref_dir: "../data/en_reference/"
output_dir: "../output/fr/"
steps: [extract, validate, align, notes]
tables: [all]
```

### For a new PDF

1. Copy `configs/_template.yaml` → `configs/bufr_XX_YYYY.yaml`
2. Set `lang`, `pdf_path`
3. Find page ranges: open the PDF, note where each table starts/ends (0-indexed)
4. Set `ref_dir` to the English reference CSVs
5. Set `output_dir`

## Step 3: Run

```bash
cd wmo_pipeline

# Dry run — check everything looks right:
python run.py --config configs/bufr_fr_2019.yaml --dry-run

# Full run:
python run.py --config configs/bufr_fr_2019.yaml
```

That's it. Wait ~10 seconds for French (larger PDFs take longer).

## Step 4: Find your output

```
output/fr/
├── raw/                  80 CSVs straight from the PDF
├── fixed/                80 CSVs with IDs validated
├── id_report.csv         what was fixed and why
├── aligned/              comparison vs English reference
├── alignment_report.md   coverage summary
└── final/                80 CSVs with noteIDs linked
```

The `final/` folder is what you want. Each CSV has:
- All text **as-is from the PDF** (never translated)
- Validated IDs matching the WMO scheme
- `noteIDs` linking to the English notes system

---

## Common tasks

### Run only specific tables

```bash
# Just Table B and CodeFlag:
python run.py -c configs/bufr_fr_2019.yaml --steps extract,validate
```

Then edit the config: `tables: [table_b, codeflag]`

### Skip extraction (re-validate existing CSVs)

```bash
python run.py -c configs/bufr_fr_2019.yaml --steps validate,align
```

(Make sure `raw/` or `fixed/` already exists from a previous run.)

### See what plugins are available

```bash
python run.py --list-plugins
```

### Calibrate page ranges for a new PDF

Open the PDF and note (0-indexed) page numbers:
- Table A usually starts after the table of contents
- Table B follows immediately after Table A
- Table C after B, Table D after C, CodeFlag at the end

Tip: search for "Table A" / "Table B" / etc. in the PDF.

---

## How the extraction works (briefly)

The pipeline reads the PDF using pymupdf's `page.get_text("dict")`,
which gives every text span with its `(x, y)` coordinates on the page.

- **x-position** → which column (FXY, name, unit, scale, etc.)
- **y-position** → which row (spans within 3pt of each other = same row)

Column boundaries are calibrated per language/PDF in `wmo/_bufr_engine/lang_config.py`.
Each language has measured `x` ranges like:

```
French Table B:
  fxy:       x = 85–150
  name:      x = 150–310
  bufr_unit: x = 310–400
  ...
```

This position-based approach handles multi-line cells, merged spans,
and complex layouts reliably.

---

## Adding a new language

1. Drop the PDF in `data/{lang}/`
2. Calibrate column positions:
   ```python
   import fitz
   doc = fitz.open("data/xx/new_pdf.pdf")
   page = doc[300]  # pick a Table B page
   for b in page.get_text("dict")["blocks"]:
       for l in b.get("lines", []):
           for s in l["spans"]:
               print(f"x={s['bbox'][0]:.0f} text={s['text'][:50]}")
   ```
3. Add a `LangConfig` in `wmo/_bufr_engine/lang_config.py`
4. Create a YAML config in `configs/`
5. Run and iterate

---

## Cross-checking with Docling

Docling uses AI models (layout analysis + table structure recognition) to
detect tables — a fundamentally different approach from pymupdf.  Run both
on the same pages to catch extraction errors.

**Install** (first run downloads ~1–2 GB of models):

```bash
uv pip install docling
```

**Quick test** (5 pages):

```bash
python tools/docling_check.py \
    --pdf data/fr/306_I2_2019_fr.pdf \
    --pages 256-260 \
    --output-dir /tmp/docling_test/
```

**Using a config** (auto-resolves page ranges):

```bash
python tools/docling_check.py \
    --config configs/bufr_fr_2019.yaml \
    --table table_b \
    --max-pages 10 \
    --output-dir /tmp/docling_test/
```

**Compare against pymupdf output**:

```bash
python tools/docling_check.py \
    --config configs/bufr_fr_2019.yaml \
    --table table_b \
    --compare-dir /path/to/output/fixed/ \
    --output-dir /tmp/docling_crosscheck/

cat /tmp/docling_crosscheck/comparison_report.md
```

Output: `docling_raw.md` (full markdown), `docling_table_N.csv` (per table),
and `comparison_report.md` (if `--compare-dir` given).

---

## File quick-reference

| File | What it does |
|------|-------------|
| `run.py` | CLI entry point |
| `configs/*.yaml` | Job configurations |
| `wmo/config.py` | Loads YAML → JobConfig |
| `wmo/runner.py` | Runs the 4-step pipeline |
| `wmo/registry.py` | Maps standards → handlers |
| `wmo/extractors/bufr.py` | BUFR table extractors |
| `wmo/validators/bufr.py` | BUFR ID validation |
| `wmo/aligners/reference.py` | Compare vs English reference |
| `wmo/notes/reference.py` | Copy noteIDs from reference |
| `wmo/_bufr_engine/` | The extraction engine |
| `wmo/_bufr_engine/lang_config.py` | Column positions per language |
| `tools/docling_check.py` | Cross-check extraction with Docling AI |
