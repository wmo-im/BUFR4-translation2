#!/usr/bin/env python3
"""
Generate PDF tables from BUFR CSV files via HTML → PDF.

Requires weasyprint: uv pip install weasyprint

Usage:
    python scripts/csv2pdf.py <lang> [--table B|D|CF|all] [--class NN] [--outdir DIR]
"""

import argparse, os, sys, glob

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from csv2docx import load_config, load_notes, find_files
from csv2html import generate_table_b_html, generate_table_d_html, generate_codeflag_html


def html_to_pdf(html_content, pdf_path):
    """Convert HTML string to PDF using weasyprint."""
    try:
        from weasyprint import HTML
    except ImportError:
        print('ERROR: weasyprint not installed. Run: uv pip install weasyprint')
        sys.exit(1)

    # Add print-specific CSS for landscape A4
    print_css = """
    <style>
      @page { size: A4 landscape; margin: 1.5cm; }
      table { page-break-inside: auto; }
      tr { page-break-inside: avoid; }
      thead { display: table-header-group; }
    </style>
    """
    html_content = html_content.replace('</head>', print_css + '</head>')
    HTML(string=html_content).write_pdf(pdf_path)


def main():
    parser = argparse.ArgumentParser(description='Generate PDF tables from BUFR CSVs')
    parser.add_argument('lang', help='Language code (ru, fr, es, en)')
    parser.add_argument('--table', default='all', choices=['B', 'D', 'CF', 'all'])
    parser.add_argument('--class', dest='class_no', default=None)
    parser.add_argument('--outdir', default=None)
    args = parser.parse_args()

    config = load_config()
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = args.outdir or os.path.join(base_dir, 'pdf', args.lang)
    os.makedirs(out_dir, exist_ok=True)

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
        pdf_path = os.path.join(out_dir, f'{basename}.pdf')
        html_to_pdf(html, pdf_path)
        total += 1
        print(f'  {basename}.pdf')

    print(f'\n{total} files written to {out_dir}')


if __name__ == '__main__':
    main()
