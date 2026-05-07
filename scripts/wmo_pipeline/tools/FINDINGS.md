# Docling Cross-Check: Findings on WMO-306 Tables

**Date:** 2026-03-17
**PDF tested:** `306_I2_2019_fr.pdf` (French, WMO-306 Vol I.2, 2019 edition)
**Pages tested:** 256–260 (Table B, Classes 00 and 01, 5 pages)
**Docling version:** installed via `uv pip install docling` (2026-03-17)

## Verdict

**Docling is not useful for WMO-306 BUFR tables.** The pymupdf position-based
extraction is dramatically better for this PDF layout.

## Why

WMO Table B has a **transposed layout** — each visual "column" in the PDF is
a data element (0 00 001, 0 00 002, ...) and each visual "row" is a field
(DESCRIPTEUR, NOM DE L'ÉLÉMENT, UNITÉ, ÉCHELLE, etc.).

Docling's AI table detector is designed for "normal" tables (rows = records,
columns = fields). It treats the WMO visual grid literally: 9 rows × N columns,
where N varies per page and multiple elements get packed into single cells.

## Specific failures

### 1. Multiple elements packed into single cells

Docling's NOM DE L'ÉLÉMENT row for page 256:

    001 Table A: entrée 002 Table A: description de la catégorie de données, ligne 1

That is **two** separate elements (000001 and 000002) concatenated into one cell.
pymupdf correctly separates them into individual rows.

### 2. FXY codes garbled across cell boundaries

Docling's DESCRIPTEUR row:

    0 00 0 00 | 0 00 003 | 0 00 004 | 0 00 005 0 | 00 006 0 00 007

`0 00 005 0` packs 000005's code with the start of 000006's.
pymupdf extracts each FXY cleanly: `000001`, `000002`, `000003`, etc.

### 3. Element names split mid-word across cells

    Cell 6: "...BUFR (voir note 2) Numéro de version de la table"
    Cell 7: "principale CREX (voir note 3) Numéro de version de la table locale"
    Cell 8: "BUFR (voir note 4) Descripteur F à ajouter..."

"table principale CREX" is split across cells 6–7. "table locale BUFR" across 7–8.
pymupdf handles these as separate complete elements.

### 4. Continuation pages collapse entirely

Page 257 (Classe 00 suite) produces only **2 columns** — all 4 remaining elements
(000025–000030) are packed into a single cell.

### 5. FXY codes corrupted on Classe 01

    0 01 001 0 01 002 0 003 0     (should be three separate FXYs)
    01 01 004 0 01 005 0          (leading "01" bleeds into prefix)

### 6. Note references merged into element names

    (voir note 1) Numéro d'édition BUFR

Note reference and next element's name concatenated in same cell.

## Numbers

| Metric | pymupdf | Docling |
|--------|---------|---------|
| Elements extracted (3 pages) | 24 (class 00) + first class 01 entries | 27 "rows" (actually field labels, not elements) |
| Tables detected | 1 continuous table per class | 3 separate tables (1 per page) |
| Column count consistency | Always 16 columns | 14, 2, and 12 columns |
| FXY codes correctly parsed | 24/24 | 0/24 (all packed or garbled) |
| Text overlap with pymupdf | — | 0.4% of cell values match |

## Why pymupdf works and Docling doesn't

pymupdf's `page.get_text("dict")` returns every text span with exact `(x, y)`
coordinates. Our extractor uses **x-position** to classify columns and
**y-position** for row alignment. This approach:

- Knows the table is transposed and handles it
- Uses calibrated column boundaries per language
- Handles multi-line cells via y-proximity grouping
- Separates note references from element names by x-position

Docling's AI model has no domain knowledge of the WMO table layout. It sees
a visual grid and tries to parse it generically — which fails on these dense,
transposed, multi-element-per-column tables.

## When Docling might be useful

Docling could still be useful for:
- **Simpler WMO tables** (Table A is a straightforward 2-column table)
- **Non-WMO PDFs** with conventional table layouts
- **Verifying text extraction** (the markdown output captures free text well,
  e.g., NOTES sections between tables)

## Raw output

The raw Docling output from this test is preserved at:
- `docling_raw.md` — full markdown rendering
- `docling_table_00.csv` through `docling_table_04.csv` — per-table CSVs
- `comparison_report.md` — automated comparison metrics

Reproduce with `docling-check --pdf <your_pdf> --pages 256-260 --output-dir /tmp/docling_test/`.
