# docling-check

Cross-check your PDF table extraction against [Docling](https://github.com/DS4SD/docling)'s AI-based table detection.

Docling uses deep learning models (layout analysis + table structure recognition) to find and parse tables in PDFs. This tool runs Docling on a PDF (or a page range), dumps the results as markdown + CSV, and optionally compares them against your own extracted CSVs — catching errors either method might miss.

## Install

```bash
pip install docling-check
```

The first run downloads Docling's AI models (~1–2 GB).

## Quick start

```bash
# Extract tables from pages 10–20:
docling-check --pdf report.pdf --pages 10-20 --output-dir /tmp/out/

# Process entire PDF:
docling-check --pdf report.pdf --output-dir /tmp/out/

# Compare against your own CSVs:
docling-check --pdf report.pdf --pages 10-20 \
    --compare-dir my_extracted_csvs/ \
    --output-dir /tmp/crosscheck/
```

## Output

```
output_dir/
├── docling_raw.md          # Full document as markdown
├── docling_table_00.csv    # Each detected table as CSV
├── docling_table_01.csv
├── ...
└── comparison_report.md    # If --compare-dir was given
```

## Usage

```
docling-check --pdf <PDF> [--pages START-END] --output-dir <DIR>
              [--compare-dir <DIR>] [--compare-glob <PATTERN>]
              [--compare-label <LABEL>] [--max-pages N]
```

### Options

| Flag | Description |
|------|-------------|
| `--pdf` | Path to source PDF |
| `--pages` | Page range, 0-indexed inclusive (e.g. `10-20`). Omit for entire PDF. |
| `--output-dir` | Directory for Docling output (required) |
| `--compare-dir` | Directory with existing CSVs to compare against |
| `--compare-glob` | Glob pattern to filter comparison CSVs (default: `*.csv`) |
| `--compare-label` | Label for the comparison source in the report (default: "existing extraction") |
| `--max-pages` | Limit pages processed (for quick tests) |

### Examples

**Basic extraction** — run Docling on 10 pages, dump markdown + CSVs:

```bash
docling-check --pdf financial_report.pdf --pages 5-14 --output-dir /tmp/docling/
```

**Cross-check** — compare Docling's tables against your Camelot/Tabula/pymupdf output:

```bash
docling-check --pdf financial_report.pdf --pages 5-14 \
    --compare-dir my_camelot_output/ \
    --compare-label "Camelot extraction" \
    --output-dir /tmp/crosscheck/

cat /tmp/crosscheck/comparison_report.md
```

**Filter comparison CSVs** — only compare against specific files:

```bash
docling-check --pdf data.pdf --pages 0-50 \
    --compare-dir output/ \
    --compare-glob "Table_B*.csv" \
    --output-dir /tmp/crosscheck/
```

**Quick test** — limit to 3 pages to verify setup:

```bash
docling-check --pdf huge_report.pdf --pages 100-200 \
    --max-pages 3 --output-dir /tmp/quick/
```

## Comparison report

When `--compare-dir` is given, the tool produces `comparison_report.md` with:

- **Row counts**: total rows from each source
- **Per-table breakdown**: rows and columns for each Docling-detected table
- **Per-file breakdown**: rows in each comparison CSV
- **Column comparison**: column names from both sources, overlap
- **Content sample**: first 10 rows from each source
- **Text content overlap**: percentage of unique cell values that appear in both outputs, with samples of values present in one but not the other

## How it works

1. **Page extraction**: Uses pymupdf to slice the requested page range into a small temp PDF (avoids processing 900-page documents through Docling)
2. **Docling conversion**: Runs Docling's `DocumentConverter` on the temp PDF
3. **Output dump**: Exports the full document as markdown and each detected table as CSV
4. **Comparison** (optional): Loads your CSVs, computes row counts, column overlap, and cell-level text overlap

## Limitations

Docling works best on PDFs with conventional table layouts (rows = records, columns = fields). It struggles with:

- **Transposed/rotated tables** where visual columns represent records
- **Dense multi-element tables** where one visual column contains multiple data items
- **Tables spanning multiple pages** with continuation headers

For such layouts, position-based extraction (pymupdf, pdfplumber) with domain-specific column calibration tends to be more reliable. See [FINDINGS.md](FINDINGS.md) for a detailed case study on WMO meteorological tables.

## Python API

```python
from docling_check import extract_pages, run_docling, docling_tables_to_dataframes, dump_output, compare_with_csvs
from pathlib import Path

# Extract 5 pages
tmp_pdf = extract_pages("report.pdf", start=10, end=14)

# Run Docling
result = run_docling(tmp_pdf)
tables = docling_tables_to_dataframes(result)

# Dump output
dump_output(result, tables, Path("/tmp/out"))

# Compare with existing CSVs
compare_with_csvs(tables, "my_csvs/", Path("/tmp/out"), label="my pipeline")

# Cleanup
tmp_pdf.unlink()
```

## License

MIT
