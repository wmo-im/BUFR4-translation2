"""
Chunk the GRIB section of the Russian WMO-306 2019 PDF into per-section PDFs.

The same 306_I2_2019_ru.pdf contains both GRIB (Part B, FM 92) and BUFR
(Part B, FM 94).  This script extracts the GRIB portion using TOC bookmarks.

Output structure:
  data/ru/chunks_grib/
    GRIB_Intro/       — Rules, octet specs
    Templates_Sec1/   — Section 1 templates
    Templates_Sec3/   — Section 3 templates (grid definition)
    Templates_Sec4/   — Section 4 templates (product definition) — largest
    Templates_Sec5/   — Section 5 templates (data representation)
    Templates_Sec7/   — Section 7 templates (data)
    CodeTable_Sec0/   — Code table for Section 0
    CodeTable_Sec1/   — Code tables for Section 1
    CodeTable_Sec3/   — Code/flag tables for Section 3
    CodeTable_Sec4/   — Code tables for Section 4 — largest
    CodeTable_Sec5/   — Code tables for Section 5
    CodeTable_Sec6/   — Code table for Section 6
    Appendix_I/       — Triangular grids
    Appendix_II/      — Arakawa grids
    Appendix_III/     — Distribution functions
    Appendix_IV/      — Time-varying tile definitions

Usage:
  cd wmo_pipeline
  python tools/chunk_pdf_grib_ru.py
"""

import os
import pymupdf

HERE     = os.path.dirname(os.path.abspath(__file__))
PIPELINE = os.path.dirname(HERE)
PDF_PATH = os.path.join(PIPELINE, 'data', 'ru', '306_I2_2019_ru.pdf')
OUT_BASE = os.path.join(PIPELINE, 'data', 'ru', 'chunks_grib')

# TOC keywords (Russian) → output folder and filename
# Each entry: (toc_keyword, folder, filename)
# Page ranges derived automatically from consecutive TOC entries.
GRIB_SECTIONS = [
    # Intro / rules
    ("FM 92",                         "GRIB_Intro",      "GRIB_Rules.pdf"),
    ("СПЕЦИФИКАЦИИ СОДЕРЖАНИЯ ОКТЕТОВ", "GRIB_Intro",      "GRIB_OctetSpecs.pdf"),
    # Templates
    ("ОБРАЗЦОВ, ИСПОЛЬЗУЕМЫХ В РАЗДЕЛЕ 1", "Templates_Sec1", "Templates_Sec1.pdf"),
    ("ОБРАЗЦОВ, ИСПОЛЬЗУЕМЫХ В РАЗДЕЛЕ 3", "Templates_Sec3", "Templates_Sec3.pdf"),
    ("ОБРАЗЦОВ, ИСПОЛЬЗУЕМЫХ В РАЗДЕЛЕ 4", "Templates_Sec4", "Templates_Sec4.pdf"),
    ("ОБРАЗЦОВ, ИСПОЛЬЗУЕМЫХ В РАЗДЕЛЕ 5", "Templates_Sec5", "Templates_Sec5.pdf"),
    ("ОБРАЗЦОВ, ИСПОЛЬЗУЕМЫХ В РАЗДЕЛЕ 7", "Templates_Sec7", "Templates_Sec7.pdf"),
    # Code/flag tables
    ("КОДОВАЯ ТАБЛИЦА, ИСПОЛЬЗУЕМАЯ В РАЗДЕЛЕ 0",           "CodeTable_Sec0", "CodeTable_Sec0.pdf"),
    ("КОДОВЫЕ ТАБЛИЦЫ, ИСПОЛЬЗУЕМЫЕ В РАЗДЕЛЕ 1",           "CodeTable_Sec1", "CodeTable_Sec1.pdf"),
    ("КОДОВЫЕ ТАБЛИЦЫ И ТАБЛИЦЫ ФЛАГОВ, ИСПОЛЬЗУЕМЫЕ В РАЗДЕЛЕ 3", "CodeTable_Sec3", "CodeTable_Sec3.pdf"),
    ("КОДОВЫЕ ТАБЛИЦЫ, ИСПОЛЬЗУЕМЫЕ В РАЗДЕЛЕ 4",           "CodeTable_Sec4", "CodeTable_Sec4.pdf"),
    ("КОДОВЫЕ ТАБЛИЦЫ, ИСПОЛЬЗУЕМЫЕ В РАЗДЕЛЕ 5",           "CodeTable_Sec5", "CodeTable_Sec5.pdf"),
    ("КОДОВАЯ ТАБЛИЦА, ИСПОЛЬЗУЕМАЯ В РАЗДЕЛЕ 6",           "CodeTable_Sec6", "CodeTable_Sec6.pdf"),
    # Appendices
    ("ДОБАВЛЕНИЕ I",   "Appendix_I",   "Appendix_I.pdf"),
    ("ДОБАВЛЕНИЕ II",  "Appendix_II",  "Appendix_II.pdf"),
    ("ДОБАВЛЕНИЕ III", "Appendix_III", "Appendix_III.pdf"),
    ("ДОБАВЛЕНИЕ IV",  "Appendix_IV",  "Appendix_IV.pdf"),
]


def save_pages(doc, p_start, p_end, out_path):
    """Save 0-indexed inclusive page range to out_path."""
    if p_end < p_start:
        print(f"  SKIP  invalid range {p_start+1}–{p_end+1}  →  {os.path.basename(out_path)}")
        return
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    new_doc = pymupdf.open()
    new_doc.insert_pdf(doc, from_page=p_start, to_page=p_end)
    new_doc.save(out_path)
    new_doc.close()
    n_pages = p_end - p_start + 1
    rel = os.path.relpath(out_path, OUT_BASE)
    print(f"  {rel:55s}  pp {p_start+1:4d}–{p_end+1:4d}  ({n_pages} pages)")


def find_toc_page(toc, keyword):
    """Return 0-indexed page of first TOC entry containing keyword."""
    for _, title, page in toc:
        if keyword in title:
            return page - 1  # convert to 0-indexed
    raise ValueError(f"TOC keyword not found: {keyword!r}")


def main():
    print(f"Opening {PDF_PATH}")
    doc = pymupdf.open(PDF_PATH)
    toc = doc.get_toc()
    print(f"  {len(toc)} TOC entries, {doc.page_count} pages total\n")

    # Resolve start pages for each section
    starts = []
    for keyword, folder, filename in GRIB_SECTIONS:
        try:
            pg = find_toc_page(toc, keyword)
            starts.append((pg, folder, filename))
        except ValueError as e:
            print(f"  WARNING: {e}")

    # Find the page where BUFR starts (FM 94) — this is the hard end of GRIB
    bufr_start = find_toc_page(toc, "FM 94")
    print(f"GRIB ends before BUFR at page {bufr_start + 1} (0-idx {bufr_start})\n")

    # Split: each section runs from its start page to the page before the next section
    total = 0
    for i, (pg_start, folder, filename) in enumerate(starts):
        if i < len(starts) - 1:
            pg_end = starts[i + 1][0] - 1
        else:
            pg_end = bufr_start - 1  # last GRIB section ends before BUFR

        out_path = os.path.join(OUT_BASE, folder, filename)
        save_pages(doc, pg_start, pg_end, out_path)
        total += 1

    doc.close()

    # Summary
    print(f"\nDone — {total} chunk files written to {OUT_BASE}")
    total_pages = sum(1 for d, _, _ in starts for _ in [1])  # just count
    print(f"\nChunk inventory:")
    for d in sorted(os.listdir(OUT_BASE)):
        full = os.path.join(OUT_BASE, d)
        if os.path.isdir(full):
            files = os.listdir(full)
            print(f"  {d:30s}  {len(files)} file(s)")


if __name__ == '__main__':
    main()
