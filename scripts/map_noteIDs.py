#!/usr/bin/env python3
"""
Build a mapping of noteIDs to the files and row IDs that reference them.

Scans all English data CSV files (TableB, TableD, CodeFlag, TableC) and
outputs a flat CSV: noteID, noteID_filename, filename, ID

Usage:
    python scripts/map_noteIDs.py [--outfile PATH]
"""

import argparse
import csv
import glob
import os

ENGLISH_DIR = 'english'
DEFAULT_OUT = os.path.join('notes', 'noteID_mapping_en.csv')

FILE_PATTERNS = [
    ('BUFRCREX_TableB_en_*.csv',  'BUFRCREX_TableB_notes_en.csv'),
    ('BUFR_TableD_en_*.csv',      'BUFR_TableD_notes_en.csv'),
    ('BUFRCREX_CodeFlag_en_*.csv', 'BUFRCREX_CodeFlag_notes_en.csv'),
    ('BUFR_TableC_en.csv',        'BUFR_TableC_notes_en.csv'),
]


def build_mapping(base_dir):
    eng_dir = os.path.join(base_dir, ENGLISH_DIR)
    rows = []

    for pattern, notes_file in FILE_PATTERNS:
        for filepath in sorted(glob.glob(os.path.join(eng_dir, pattern))):
            filename = os.path.basename(filepath)
            with open(filepath, encoding='utf-8-sig') as f:
                for row in csv.DictReader(f):
                    note_ids = row.get('noteIDs', '').strip()
                    row_id = row.get('ID', '').strip()
                    if not note_ids or not row_id:
                        continue
                    for nid in note_ids.split(','):
                        nid = nid.strip()
                        if nid:
                            rows.append((nid, notes_file, filename, row_id))

    rows.sort(key=lambda r: (int(r[0]) if r[0].isdigit() else float('inf'), r[2], r[3]))
    return rows


def main():
    parser = argparse.ArgumentParser(description='Map noteIDs to file references')
    parser.add_argument('--outfile', default=None)
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_path = args.outfile or os.path.join(base_dir, DEFAULT_OUT)

    rows = build_mapping(base_dir)

    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['noteID', 'noteID_filename', 'filename', 'ID'])
        writer.writerows(rows)

    print(f'{len(rows)} references written to {out_path}')


if __name__ == '__main__':
    main()
