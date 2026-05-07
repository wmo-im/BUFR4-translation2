#!/usr/bin/env python3
"""
Generate HTML tables from BUFR CSV files.

Uses table_config.json for language-specific titles, headers, and styling.
Shares the same config as csv2docx.py.

Usage:
    python scripts/csv2html.py <lang> [--table B|D|CF|all] [--class NN] [--outdir DIR]
"""

import argparse, csv, json, os, sys, glob
from html import escape

# Reuse config/notes/FXY logic from csv2docx
from csv2docx import load_config, get_lang_config, load_notes, resolve_note, format_fxy, get_row_values


CSS = """
<style>
  body { font-family: Verdana, sans-serif; font-size: 8pt; margin: 1.5cm; }
  h1 { font-size: 11pt; text-align: center; }
  h2 { font-size: 9pt; margin-top: 1em; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 1em; }
  th { background-color: #D9D9D9; font-weight: normal; font-size: 8pt;
       text-align: left; padding: 4px 6px; border: 1px solid #595959; }
  td { font-size: 8pt; padding: 4px 6px; border: 1px solid #595959;
       vertical-align: top; }
  tr.seq-header td { font-weight: bold; }
  .fxy-heading { font-weight: bold; font-size: 9pt; margin: 0.8em 0 0.3em 0; }
</style>
"""


def html_header(title):
    return f'<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>{escape(title)}</title>\n{CSS}\n</head>\n<body>\n<h1>{escape(title)}</h1>\n'


def html_footer():
    return '</body>\n</html>\n'


def html_table_start(headers):
    hdr = ''.join(f'<th>{escape(h)}</th>' for h in headers)
    return f'<table>\n<thead><tr>{hdr}</tr></thead>\n<tbody>\n'


def html_table_end():
    return '</tbody>\n</table>\n'


def html_row(vals, css_class=''):
    cls = f' class="{css_class}"' if css_class else ''
    cells = ''.join(f'<td>{escape(str(v or ""))}</td>' for v in vals)
    return f'<tr{cls}>{cells}</tr>\n'


# ── Table B ──────────────────────────────────────────────────────────

def generate_table_b_html(lang, csv_path, config, notes_db):
    lc = get_lang_config(config, lang, 'table_b')
    headers = lc.get('headers', [])
    col_map = lc.get('column_map', [])

    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return ''

    class_no = rows[0].get('ClassNo', '')
    class_name = rows[0].get(f'ClassName_{lang}', '')
    title = lc.get('title', '').format(class_no=class_no, class_name=class_name)

    fxy_fmt = lc.get('fxy_format', 'compact')
    out = html_header(title) + html_table_start(headers)
    for r in rows:
        vals = get_row_values(r, col_map, lang, notes_db, fxy_fmt)
        out += html_row(vals)
    out += html_table_end() + html_footer()
    return out


# ── Table D ──────────────────────────────────────────────────────────

def generate_table_d_html(lang, csv_path, config, notes_db):
    lc = get_lang_config(config, lang, 'table_d')
    headers = lc.get('headers', [])
    col_map = lc.get('column_map', [])

    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return ''

    cat = rows[0].get('Category', '')
    cat_name = rows[0].get(f'CategoryOfSequences_{lang}', '')
    title = lc.get('title', '').format(category=cat, category_name=cat_name)

    fxy_fmt = lc.get('fxy_format', 'spaced')
    out = html_header(title) + html_table_start(headers)
    prev_fxy1 = None
    for r in rows:
        fxy1 = r.get('FXY1', '')
        vals = get_row_values(r, col_map, lang, notes_db, fxy_fmt)
        is_new_seq = fxy1 != prev_fxy1
        if not is_new_seq:
            vals[0] = ''
            vals[1] = ''
        prev_fxy1 = fxy1
        out += html_row(vals, 'seq-header' if is_new_seq else '')
    out += html_table_end() + html_footer()
    return out


# ── CodeFlag ─────────────────────────────────────────────────────────

def generate_codeflag_html(lang, csv_path, config, notes_db):
    lc = get_lang_config(config, lang, 'codeflag')

    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return ''

    fxy_groups = {}
    for r in rows:
        fxy_groups.setdefault(r.get('FXY', ''), []).append(r)

    cls = os.path.basename(csv_path).split('_')[-1].replace('.csv', '')
    title = lc.get('title', '').format(class_no=cls)
    out = html_header(title)

    for fxy, group in fxy_groups.items():
        elem = group[0].get(f'ElementName_{lang}', '')
        fxy_fmt = lc.get('fxy_format', 'compact')
        out += f'<p class="fxy-heading">{escape(format_fxy(fxy, fxy_fmt))}  {escape(elem)}</p>\n'

        has_sub1 = any(r.get(f'EntryName_sub1_{lang}', '').strip() for r in group)
        has_sub2 = any(r.get(f'EntryName_sub2_{lang}', '').strip() for r in group)

        if has_sub2:
            headers = lc.get('headers_sub2', lc['headers'])
            col_map = lc.get('column_map_sub2', lc['column_map'])
        elif has_sub1:
            headers = lc.get('headers_sub1', lc['headers'])
            col_map = lc.get('column_map_sub1', lc['column_map'])
        else:
            headers = lc['headers']
            col_map = lc['column_map']

        out += html_table_start(headers)
        for r in group:
            vals = get_row_values(r, col_map, lang, notes_db, fxy_fmt)
            out += html_row(vals)
        out += html_table_end()

    out += html_footer()
    return out


# ── Main ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Generate HTML tables from BUFR CSVs')
    parser.add_argument('lang', help='Language code (ru, fr, es, en)')
    parser.add_argument('--table', default='all', choices=['B', 'D', 'CF', 'all'])
    parser.add_argument('--class', dest='class_no', default=None)
    parser.add_argument('--outdir', default=None)
    args = parser.parse_args()

    # Add scripts dir to path for csv2docx imports
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

    config = load_config()
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = args.outdir or os.path.join(base_dir, 'html', args.lang)
    os.makedirs(out_dir, exist_ok=True)

    from csv2docx import find_files
    files = find_files(args.lang, args.table, args.class_no, base_dir)
    if not files:
        print(f'No files found')
        sys.exit(1)

    notes_db = load_notes(args.lang, base_dir)
    print(f'Loaded {len(notes_db)} notes for {args.lang}')

    generators = {
        'B': generate_table_b_html,
        'D': generate_table_d_html,
        'CF': generate_codeflag_html,
    }

    total = 0
    for table_type, csv_path in files:
        html = generators[table_type](args.lang, csv_path, config, notes_db)
        basename = os.path.splitext(os.path.basename(csv_path))[0]
        out_path = os.path.join(out_dir, f'{basename}.html')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html)
        total += 1
        print(f'  {basename}.html')

    print(f'\n{total} files written to {out_dir}')


if __name__ == '__main__':
    main()
