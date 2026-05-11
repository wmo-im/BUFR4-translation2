#!/usr/bin/env python3
"""
Add table-type prefixes to all noteIDs across notes and data files.

Prefixes applied:
    BUFRCREX_TableB   -> B   (e.g. 11 -> B11)
    BUFR_TableD       -> D   (e.g. 11 -> D11)
    BUFRCREX_CodeFlag -> CF  (e.g. 11 -> CF11)
    BUFR_TableC       -> C   (e.g. 11 -> C11)

Files updated:
    notes/*_notes_{lang}.csv        -- noteID column
    notes/*_global_noteIDs.csv      -- noteID column
    english|french|spanish|russian  -- noteIDs column (comma-separated)

Usage:
    python scripts/prefix_noteIDs.py [--dry-run]
"""

import argparse
import csv
import glob
import io
import os

LANG_DIRS = {
    'en': 'english',
    'fr': 'french',
    'es': 'spanish',
    'ru': 'russian',
}

# (data file glob pattern, notes file prefix, noteID prefix)
TABLE_TYPES = [
    ('BUFRCREX_TableB',   'B'),
    ('BUFR_TableD',       'D'),
    ('BUFRCREX_CodeFlag', 'CF'),
    ('BUFR_TableC',       'C'),
]


def add_prefix(value, prefix):
    """Prepend prefix to a noteID if not already prefixed."""
    v = value.strip()
    if not v or not v[0].isdigit():
        return v  # already prefixed or empty
    return prefix + v


def prefix_noteid_column(filepath, prefix, col, dry_run):
    """Update a single noteID column value (not comma-separated) in a CSV."""
    with open(filepath, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        rows = list(reader)

    if col not in fieldnames:
        return 0

    changed = 0
    for row in rows:
        old = row[col]
        new = add_prefix(old, prefix)
        if new != old:
            row[col] = new
            changed += 1

    if changed and not dry_run:
        with open(filepath, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

    return changed


def prefix_noteids_column(filepath, prefix, dry_run):
    """Update comma-separated noteIDs column in a data file."""
    with open(filepath, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        rows = list(reader)

    if 'noteIDs' not in fieldnames:
        return 0

    changed = 0
    for row in rows:
        old = row['noteIDs'].strip()
        if not old:
            continue
        parts = [add_prefix(p, prefix) for p in old.split(',')]
        new = ','.join(parts)
        if new != old:
            row['noteIDs'] = new
            changed += 1

    if changed and not dry_run:
        with open(filepath, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

    return changed


def main():
    parser = argparse.ArgumentParser(description='Prefix noteIDs by table type')
    parser.add_argument('--dry-run', action='store_true',
                        help='Report changes without writing files')
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    notes_dir = os.path.join(base_dir, 'notes')
    total = 0

    for table_prefix, prefix in TABLE_TYPES:
        # 1. Notes files for all languages
        for notes_file in sorted(glob.glob(
                os.path.join(notes_dir, f'{table_prefix}_notes_*.csv'))):
            n = prefix_noteid_column(notes_file, prefix, 'noteID', args.dry_run)
            if n:
                print(f'  {os.path.basename(notes_file)}: {n} noteIDs prefixed')
                total += n

        # 2. Global noteIDs file
        global_file = os.path.join(notes_dir, f'{table_prefix}_global_noteIDs.csv')
        if os.path.exists(global_file):
            n = prefix_noteid_column(global_file, prefix, 'noteID', args.dry_run)
            if n:
                print(f'  {os.path.basename(global_file)}: {n} noteIDs prefixed')
                total += n

        # 3. Data files for all languages
        for lang, lang_dir_name in LANG_DIRS.items():
            lang_dir = os.path.join(base_dir, lang_dir_name)
            pattern = os.path.join(lang_dir, f'{table_prefix}_{lang}*.csv')
            for data_file in sorted(glob.glob(pattern)):
                n = prefix_noteids_column(data_file, prefix, args.dry_run)
                if n:
                    print(f'  {os.path.basename(data_file)}: {n} rows updated')
                    total += n

    action = 'Would update' if args.dry_run else 'Updated'
    print(f'\n{action} {total} values total')


if __name__ == '__main__':
    main()
