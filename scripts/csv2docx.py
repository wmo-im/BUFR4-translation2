#!/usr/bin/env python3
"""
Generate Word (.docx) tables from BUFR CSV files.

Produces Word documents matching PDF content for Table B, Table D, and CodeFlag.

Usage:
    python scripts/csv2docx.py <lang> [--table B|D|CF|all] [--class NN] [--outdir DIR]

Examples:
    python scripts/csv2docx.py ru                          # all tables, all classes
    python scripts/csv2docx.py ru --table B                # all Table B classes
    python scripts/csv2docx.py ru --table D --class 07     # Table D class 07 only
    python scripts/csv2docx.py fr --table CF --class 20    # French CodeFlag class 20
    python scripts/csv2docx.py ru --outdir /tmp/docx       # custom output directory
"""

import argparse
import csv
import os
import sys
import glob

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT


# ── Notes loader ─────────────────────────────────────────────────────

def load_notes(lang, base_dir):
    """Load all notes files for a language. Returns dict: noteID → note text.
    Falls back to English if the language-specific note doesn't exist."""
    notes_dir = os.path.join(base_dir, 'notes')
    all_notes = {}

    # Table type prefixes for notes files
    prefixes = ['BUFRCREX_CodeFlag', 'BUFRCREX_TableB', 'BUFR_TableC', 'BUFR_TableD']

    for prefix in prefixes:
        # Only load notes for the requested language — NO English fallback
        notes_file = os.path.join(notes_dir, f'{prefix}_notes_{lang}.csv')
        if os.path.exists(notes_file):
            with open(notes_file, encoding='utf-8-sig') as f:
                for r in csv.DictReader(f):
                    nid = r.get('noteID', '').strip()
                    note_col = f'note_{lang}'
                    text = r.get(note_col, '').strip()
                    if not text:
                        text = r.get('note', '').strip()
                    if nid and text and nid not in all_notes:
                        all_notes[nid] = text

    return all_notes


def resolve_note(row, lang, notes_db):
    """Resolve the Note column: if noteIDs is populated, return full note text."""
    note_text = row.get(f'Note_{lang}', '').strip()
    note_ids = row.get('noteIDs', '').strip()

    if not note_ids:
        return note_text

    # noteIDs can be comma-separated
    resolved = []
    for nid in note_ids.split(','):
        nid = nid.strip()
        if nid in notes_db:
            resolved.append(f'Note {nid}: {notes_db[nid]}')

    if resolved:
        return ' | '.join(resolved)
    # No matching notes found — keep the original short reference as-is
    return note_text


# ── Styling ──────────────────────────────────────────────────────────

FONT_NAME = 'Arial'
FONT_SIZE_TITLE = Pt(11)
FONT_SIZE_HEADER = Pt(8)
FONT_SIZE_DATA = Pt(8)
FONT_SIZE_NOTE = Pt(7)


def set_cell_font(cell, text, size=FONT_SIZE_DATA, bold=False):
    """Set cell text with consistent font."""
    cell.text = ''
    p = cell.paragraphs[0]
    p.space_before = Pt(1)
    p.space_after = Pt(1)
    run = p.add_run(str(text) if text else '')
    run.font.name = FONT_NAME
    run.font.size = size
    run.font.bold = bold
    return run


def add_header_row(table, headers, widths=None):
    """Style the header row."""
    row = table.rows[0]
    for i, (cell, hdr) in enumerate(zip(row.cells, headers)):
        set_cell_font(cell, hdr, size=FONT_SIZE_HEADER, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Light gray background
        from docx.oxml.ns import qn
        shading = cell._element.get_or_add_tcPr()
        sh = shading.makeelement(qn('w:shd'), {
            qn('w:val'): 'clear',
            qn('w:color'): 'auto',
            qn('w:fill'): 'D9E2F3',
        })
        shading.append(sh)


def set_landscape(doc):
    """Set page to landscape A4."""
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)


def add_title(doc, text):
    """Add document title."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_TITLE
    run.font.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ── Table B Generator ────────────────────────────────────────────────

def generate_table_b(lang, csv_path, doc, notes_db=None):
    """Generate Table B Word table from CSV."""
    notes_db = notes_db or {}
    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))

    if not rows:
        return

    # Extract class info from first row
    class_no = rows[0].get('ClassNo', '')
    class_name = rows[0].get(f'ClassName_{lang}', '')
    add_title(doc, f'BUFR/CREX Table B — Class {class_no}: {class_name}')

    headers = ['FXY', f'ElementName', f'BUFR Unit', 'Scale', 'Ref Value', 'Bits',
               f'CREX Unit', 'Scale', 'Chars', 'Note']
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_header_row(table, headers)

    for r in rows:
        row_cells = table.add_row().cells
        vals = [
            r.get('FXY', ''),
            r.get(f'ElementName_{lang}', ''),
            r.get(f'BUFR_Unit_{lang}', ''),
            r.get('BUFR_Scale', ''),
            r.get('BUFR_ReferenceValue', ''),
            r.get('BUFR_DataWidth_Bits', ''),
            r.get(f'CREX_Unit_{lang}', ''),
            r.get('CREX_Scale', ''),
            r.get('CREX_DataWidth_Char', ''),
            resolve_note(r, lang, notes_db),
        ]
        for cell, val in zip(row_cells, vals):
            set_cell_font(cell, val)


# ── Table D Generator ────────────────────────────────────────────────

def generate_table_d(lang, csv_path, doc, notes_db=None):
    """Generate Table D Word table from CSV."""
    notes_db = notes_db or {}
    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))

    if not rows:
        return

    cat = rows[0].get('Category', '')
    cat_name = rows[0].get(f'CategoryOfSequences_{lang}', '')
    add_title(doc, f'BUFR Table D — Category {cat}: {cat_name}')

    headers = ['FXY1', f'Title', f'SubTitle', 'FXY2',
               f'ElementName', f'Description', 'Note']
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_header_row(table, headers)

    prev_fxy1 = None
    for r in rows:
        row_cells = table.add_row().cells
        fxy1 = r.get('FXY1', '')
        # Only show FXY1/Title on first row of each sequence
        show_fxy1 = fxy1 if fxy1 != prev_fxy1 else ''
        show_title = r.get(f'Title_{lang}', '') if fxy1 != prev_fxy1 else ''
        prev_fxy1 = fxy1

        vals = [
            show_fxy1,
            show_title,
            r.get(f'SubTitle_{lang}', ''),
            r.get('FXY2', ''),
            r.get(f'ElementName_{lang}', ''),
            r.get(f'ElementDescription_{lang}', ''),
            resolve_note(r, lang, notes_db),
        ]
        for cell, val in zip(row_cells, vals):
            set_cell_font(cell, val)

        # Bold the FXY1 row
        if show_fxy1:
            for cell in row_cells[:2]:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.bold = True


# ── CodeFlag Generator ───────────────────────────────────────────────

def generate_codeflag(lang, csv_path, doc, notes_db=None):
    """Generate CodeFlag Word table from CSV."""
    notes_db = notes_db or {}
    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))

    if not rows:
        return

    # Group by FXY
    fxy_groups = {}
    for r in rows:
        fxy = r.get('FXY', '')
        fxy_groups.setdefault(fxy, []).append(r)

    cls = os.path.basename(csv_path).split('_')[-1].replace('.csv', '')
    add_title(doc, f'BUFR/CREX Code/Flag Tables — Class {cls}')

    for fxy, group in fxy_groups.items():
        elem_name = group[0].get(f'ElementName_{lang}', '')

        # Sub-heading for each FXY
        p = doc.add_paragraph()
        run = p.add_run(f'{fxy}  {elem_name}')
        run.font.name = FONT_NAME
        run.font.size = Pt(9)
        run.font.bold = True

        has_sub1 = any(r.get(f'EntryName_sub1_{lang}', '').strip() for r in group)
        has_sub2 = any(r.get(f'EntryName_sub2_{lang}', '').strip() for r in group)

        headers = ['Code Figure', f'Entry Name']
        if has_sub1:
            headers.append('Sub 1')
        if has_sub2:
            headers.append('Sub 2')
        headers.append('Note')

        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        add_header_row(table, headers)

        for r in group:
            row_cells = table.add_row().cells
            vals = [
                r.get('CodeFigure', ''),
                r.get(f'EntryName_{lang}', ''),
            ]
            if has_sub1:
                vals.append(r.get(f'EntryName_sub1_{lang}', ''))
            if has_sub2:
                vals.append(r.get(f'EntryName_sub2_{lang}', ''))
            vals.append(resolve_note(r, lang, notes_db))

            for cell, val in zip(row_cells, vals):
                set_cell_font(cell, val)

        doc.add_paragraph()  # spacing between FXY tables


# ── Main ─────────────────────────────────────────────────────────────

def find_files(lang, table_type, class_no, base_dir):
    """Find CSV files matching the criteria."""
    lang_dir = os.path.join(base_dir, {
        'en': 'english', 'fr': 'french', 'es': 'spanish', 'ru': 'russian'
    }.get(lang, lang))

    patterns = {
        'B': f'BUFRCREX_TableB_{lang}_*.csv',
        'D': f'BUFR_TableD_{lang}_*.csv',
        'CF': f'BUFRCREX_CodeFlag_{lang}_*.csv',
    }

    if table_type == 'all':
        types = ['B', 'D', 'CF']
    else:
        types = [table_type]

    files = []
    for t in types:
        matched = sorted(glob.glob(os.path.join(lang_dir, patterns[t])))
        if class_no:
            matched = [f for f in matched if f'_{class_no}.' in f or f'_{class_no.zfill(2)}.' in f]
        files.extend([(t, f) for f in matched])

    return files


def main():
    parser = argparse.ArgumentParser(description='Generate Word tables from BUFR CSVs')
    parser.add_argument('lang', help='Language code (ru, fr, es, en)')
    parser.add_argument('--table', default='all', choices=['B', 'D', 'CF', 'all'],
                        help='Table type: B, D, CF, or all')
    parser.add_argument('--class', dest='class_no', default=None,
                        help='Class number (e.g., 01, 07)')
    parser.add_argument('--outdir', default=None,
                        help='Output directory (default: <repo>/docx/<lang>/)')
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = args.outdir or os.path.join(base_dir, 'docx', args.lang)
    os.makedirs(out_dir, exist_ok=True)

    files = find_files(args.lang, args.table, args.class_no, base_dir)
    if not files:
        print(f'No files found for lang={args.lang} table={args.table} class={args.class_no}')
        sys.exit(1)

    # Load notes database for this language
    notes_db = load_notes(args.lang, base_dir)
    print(f'Loaded {len(notes_db)} notes for {args.lang}')

    generators = {
        'B': generate_table_b,
        'D': generate_table_d,
        'CF': generate_codeflag,
    }

    total = 0
    for table_type, csv_path in files:
        doc = Document()
        set_landscape(doc)

        basename = os.path.splitext(os.path.basename(csv_path))[0]
        generators[table_type](args.lang, csv_path, doc, notes_db)

        out_path = os.path.join(out_dir, f'{basename}.docx')
        doc.save(out_path)
        total += 1
        print(f'  {basename}.docx')

    print(f'\n{total} files written to {out_dir}')


if __name__ == '__main__':
    main()
