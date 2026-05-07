"""
Chunk the Russian WMO-306 2019 PDF into per-table / per-class PDFs.

Outer page boundaries come from configs/bufr_ru_2019.yaml (matching the
pipeline's extraction ranges exactly).  Inner class/category splits come
from the PDF's table-of-contents.

Output structure mirrors data/fr/chunks/:
  data/ru/chunks/
    TableA/TableA.pdf
    TableB/TableB_intro.pdf, TableB_00.pdf, TableB_01.pdf, ...
    TableC/TableC.pdf
    TableD/TableD_intro.pdf, TableD_00.pdf, TableD_01.pdf, ...
    CodeFlag/CodeFlag_intro.pdf, CodeFlag_01.pdf, ...

Usage:
  cd wmo_pipeline
  python tools/chunk_pdf_ru.py
"""

import os
import re
import pymupdf

HERE     = os.path.dirname(os.path.abspath(__file__))
PIPELINE = os.path.dirname(HERE)
PDF_PATH = os.path.join(PIPELINE, 'data', 'ru', '306_I2_2019_ru.pdf')
OUT_BASE = os.path.join(PIPELINE, 'data', 'ru', 'chunks')

# Page ranges from configs/bufr_ru_2019.yaml  (0-indexed, inclusive)
PAGE_RANGES = {
    'a':        (248, 249),
    'b':        (251, 379),
    'c':        (380, 384),
    'd':        (386, 626),
    'codeflag': (627, 830),
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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
    rel = os.path.relpath(out_path, OUT_BASE)
    print(f"  {rel:55s}  (PDF pp {p_start+1}–{p_end+1})")


def find_toc_idx(toc, keyword):
    """Return index of first TOC entry whose title contains keyword."""
    for i, (_, title, _) in enumerate(toc):
        if keyword in title:
            return i
    raise ValueError(f"TOC keyword not found: {keyword!r}")


def collect_children(toc, parent_idx, pattern, min_page=10):
    """
    Collect TOC child entries matching `pattern` (group 1 = number).
    Skips entries with page <= min_page (broken bookmark targets).
    Returns list of dicts: {num, page, toc_i}
    """
    parent_level = toc[parent_idx][0]
    children = []
    for i in range(parent_idx + 1, len(toc)):
        level, title, page = toc[i]
        if level <= parent_level:
            break
        m = re.search(pattern, title)
        if m and page > min_page:
            children.append({'num': m.group(1), 'page': page, 'toc_i': i})
    return children


def collect_numeric_children(toc, parent_idx, min_page=10):
    """
    Collect TOC child entries whose entire title is a number (e.g. '01', '02').
    Used for CodeFlag sub-sections in the Russian PDF.
    """
    parent_level = toc[parent_idx][0]
    children = []
    for i in range(parent_idx + 1, len(toc)):
        level, title, page = toc[i]
        if level <= parent_level:
            break
        if re.fullmatch(r'\d+', title.strip()) and page > min_page:
            children.append({'num': title.strip(), 'page': page, 'toc_i': i})
    return children


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print(f"Opening {PDF_PATH}")
    doc = pymupdf.open(PDF_PATH)
    toc = doc.get_toc()

    # Locate TOC anchors needed for inner class/category splits
    idx_b  = find_toc_idx(toc, "Таблица В кодов BUFR/CREX")
    idx_d  = find_toc_idx(toc, "Таблица D кода BUFR")
    idx_cf = find_toc_idx(toc, "КОДОВЫЕ ТАБЛИЦЫ И ТАБЛИЦЫ ФЛАГОВ, СВЯЗАННЫЕ С ТАБЛИЦЕЙ В")

    # Find the appendix (ДОБАВЛЕНИЕ) that follows CodeFlag — used as the true
    # end of CodeFlag content (the config range may be 1 page short of the last class).
    cf_level = toc[idx_cf][0]
    idx_cf_end = len(toc) - 1
    for i in range(idx_cf + 1, len(toc)):
        if toc[i][0] <= cf_level:
            idx_cf_end = i
            break
    codeflag_last_page = toc[idx_cf_end][2] - 2   # 0-indexed last page before appendix

    pr = PAGE_RANGES   # shorthand

    # ========= Table A =====================================================
    print("\n=== Table A ===")
    save_pages(doc, *pr['a'], os.path.join(OUT_BASE, 'TableA', 'TableA.pdf'))

    # ========= Table B =====================================================
    print("\n=== Table B ===")
    b_classes = collect_children(toc, idx_b, r'Класс\s+(\d+)')
    print(f"  Classes found: {[c['num'] for c in b_classes]}")

    # Intro: Table B start … first class - 1
    save_pages(
        doc,
        pr['b'][0],
        b_classes[0]['page'] - 2,           # one page before first class
        os.path.join(OUT_BASE, 'TableB', 'TableB_intro.pdf'),
    )
    # Per-class chunks (inner boundaries from TOC; last class ends at Table B end)
    for i, cls in enumerate(b_classes):
        ep = (b_classes[i+1]['page'] - 2) if i < len(b_classes) - 1 else pr['b'][1]
        save_pages(
            doc,
            cls['page'] - 1,
            ep,
            os.path.join(OUT_BASE, 'TableB', f"TableB_{cls['num'].zfill(2)}.pdf"),
        )

    # ========= Table C =====================================================
    print("\n=== Table C ===")
    save_pages(doc, *pr['c'], os.path.join(OUT_BASE, 'TableC', 'TableC.pdf'))

    # ========= Table D =====================================================
    print("\n=== Table D ===")
    d_cats = collect_children(toc, idx_d, r'Категория\s+(\d+)')
    print(f"  Categories found: {[c['num'] for c in d_cats]}")

    # Intro
    save_pages(
        doc,
        pr['d'][0],
        d_cats[0]['page'] - 2,
        os.path.join(OUT_BASE, 'TableD', 'TableD_intro.pdf'),
    )
    # Per-category chunks
    for i, cat in enumerate(d_cats):
        ep = (d_cats[i+1]['page'] - 2) if i < len(d_cats) - 1 else pr['d'][1]
        save_pages(
            doc,
            cat['page'] - 1,
            ep,
            os.path.join(OUT_BASE, 'TableD', f"TableD_{cat['num'].zfill(2)}.pdf"),
        )

    # ========= CodeFlag ====================================================
    print("\n=== CodeFlag ===")
    cf_subs = collect_numeric_children(toc, idx_cf)
    print(f"  Sub-sections found: {[c['num'] for c in cf_subs]}")

    # Intro
    save_pages(
        doc,
        pr['codeflag'][0],
        cf_subs[0]['page'] - 2,
        os.path.join(OUT_BASE, 'CodeFlag', 'CodeFlag_intro.pdf'),
    )
    # Per-sub-section chunks
    for i, sub in enumerate(cf_subs):
        ep = (cf_subs[i+1]['page'] - 2) if i < len(cf_subs) - 1 else codeflag_last_page
        save_pages(
            doc,
            sub['page'] - 1,
            ep,
            os.path.join(OUT_BASE, 'CodeFlag', f"CodeFlag_{sub['num'].zfill(2)}.pdf"),
        )

    doc.close()

    # ---- summary -----------------------------------------------------------
    total = sum(
        len(os.listdir(os.path.join(OUT_BASE, d)))
        for d in ('TableA', 'TableB', 'TableC', 'TableD', 'CodeFlag')
        if os.path.isdir(os.path.join(OUT_BASE, d))
    )
    print(f"\nDone — {total} chunk files written to {OUT_BASE}")


if __name__ == '__main__':
    main()
