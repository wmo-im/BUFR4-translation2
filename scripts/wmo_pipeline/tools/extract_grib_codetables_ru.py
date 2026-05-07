#!/usr/bin/env python3
"""
Extract GRIB2 code/flag tables from Russian WMO-306 2019 PDF chunks.

Reads:  data/ru/chunks_grib/CodeTable_Sec{0,1,3,4,5,6}/*.pdf
Writes: data/ru/grib_output/codeflags/*.csv

Handles five layout types:
  1. Standard CodeTable  (Sec 0,1, most of 3+4, 5, 6) — 2-column
  2. FlagTable           (some Sec 3)                  — 3-column (bit/val/meaning)
  3. Table 4.1           (categories by discipline)    — shifted columns, single file
  4. Table 4.2           (params by disc+cat)          — 3-column with units, split files
  5. CodeTable with Notes (some Sec 3)                 — 2-column + notes col

Usage:
    cd wmo_pipeline
    python tools/extract_grib_codetables_ru.py          # all sections
    python tools/extract_grib_codetables_ru.py 0 1      # specific sections
"""

import os, re, csv, sys
from collections import OrderedDict
import pymupdf

# ── Paths ──────────────────────────────────────────────────────────────
HERE     = os.path.dirname(os.path.abspath(__file__))
PIPELINE = os.path.dirname(HERE)
CHUNKS   = os.path.join(PIPELINE, 'data', 'ru', 'chunks_grib')
OUTDIR   = os.path.join(PIPELINE, 'data', 'ru', 'grib_output', 'codeflags')

# ── Layout constants (from PDF span analysis) ─────────────────────────
Y_MIN, Y_MAX = 55, 750          # content band (excl. header at y~41, footer at y~780)
Y_TOL        = 3.0              # same-line y-tolerance

TITLE_SZ     = 9.4              # table title / sub-header font size threshold
DATA_SZ_LO   = 8.3              # data row font size range
DATA_SZ_HI   = 9.2

# Standard CodeTable columns
STD_CODE_X   = 140              # code-figure spans: x < 140
STD_VAL_X    = 143              # value/description spans: x >= 143

# FlagTable columns
FLG_BIT_X    = 138              # bit number: x < 138
FLG_MEAN_X   = 190              # meaning text: x >= 190  (value 0/1 between)

# Table 4.1 / 4.2 columns (shifted right)
T4_DESC_X    = 220              # number: x < 220, description: x >= 220
T4_UNITS_X   = 430              # units column: x >= 430

# Optional notes column in some Sec 3 tables
NOTES_X      = 315

# ── Regex ──────────────────────────────────────────────────────────────
RE_CT   = re.compile(r'Кодовая\s+таблица\s+(\d+)\.(\d+)')
RE_FT   = re.compile(r'Таблица\s+флагов\s+(\d+)\.(\d+)')
RE_DISC = re.compile(r'Дисциплина\s+продукции\s+(\d+)')
RE_CAT  = re.compile(r'(?:категория\s+)?параметра\s+(\d+)')
RE_CONT = re.compile(r'продолж\.')
RE_NUM  = re.compile(r'^[\d]+(?:[\-–—][\d]+)?$')

CSV_COLS = ['ID', 'Title_ru', 'SubTitle_ru', 'CodeFlag', 'Value',
            'MeaningParameterDescription_ru', 'Note_ru', 'noteIDs',
            'UnitComments_ru', 'Status', 'translated_by']


# ── Helpers ────────────────────────────────────────────────────────────

def norm(s):
    """Normalize en-dash / em-dash → hyphen."""
    return s.replace('\u2013', '-').replace('\u2014', '-')

def is_num(text):
    return bool(RE_NUM.match(norm(text.strip().replace(' ', ''))))

def append_text(base, new):
    """Append text, rejoining words broken by PDF line-break hyphens."""
    if not new:
        return base
    if not base:
        return new
    # If previous text ends with hyphen and new starts lowercase → rejoin word
    if base.endswith('-') and new and new[0].islower():
        return base[:-1] + new
    return base + ' ' + new

def is_title(sp):
    return sp['size'] >= TITLE_SZ and 'Bold' in sp['font'] and 'Semi' not in sp['font']

def is_subhd(sp):
    return sp['size'] >= TITLE_SZ and 'SemiBold' in sp['font']

def is_data(sp):
    return DATA_SZ_LO <= sp['size'] <= DATA_SZ_HI

def is_small(sp):
    return sp['size'] < DATA_SZ_LO


def get_spans(page):
    """Extract content spans from one PDF page."""
    out = []
    for blk in page.get_text("dict")["blocks"]:
        if "lines" not in blk:
            continue
        for line in blk["lines"]:
            for sp in line["spans"]:
                y = sp["bbox"][1]
                if y < Y_MIN or y > Y_MAX:
                    continue
                t = sp["text"].strip()
                if not t:
                    continue
                out.append({'x': sp["bbox"][0], 'y': sp["bbox"][1],
                            'text': t, 'font': sp["font"],
                            'size': round(sp["size"], 1)})
    return out


def group_y(spans):
    """Group spans into lines by y-proximity."""
    if not spans:
        return []
    ss = sorted(spans, key=lambda s: (s['y'], s['x']))
    lines = [[ss[0]]]
    for s in ss[1:]:
        if abs(s['y'] - lines[-1][0]['y']) <= Y_TOL:
            lines[-1].append(s)
        else:
            lines.append([s])
    return [sorted(ln, key=lambda s: s['x']) for ln in lines]


# ── Phase 1: Walk PDF pages, detect table boundaries ──────────────────

def walk_chunk(sec_num):
    """Walk a chunk PDF page-by-page. Returns list of table dicts."""
    chunk_dir = os.path.join(CHUNKS, f'CodeTable_Sec{sec_num}')
    pdfs = sorted(f for f in os.listdir(chunk_dir) if f.endswith('.pdf'))
    doc = pymupdf.open(os.path.join(chunk_dir, pdfs[0]))

    tables = []
    cur = None
    cur_disc = cur_cat = None
    cur_sub = ''
    subhdr_buf = []

    for pg in range(len(doc)):
        for line in group_y(get_spans(doc[pg])):
            full = ' '.join(s['text'] for s in line)

            # ── continuation markers → skip
            if RE_CONT.search(full) and all(is_small(s) or 'Itali' in s['font'] for s in line):
                continue

            # ── table title (Bold, ≥9.5pt)
            if any(is_title(s) for s in line):
                m = RE_CT.search(full)
                if not m:
                    m = RE_FT.search(full)
                if m:
                    s_n, t_n = int(m.group(1)), int(m.group(2))
                    ttype = 'FlagTable' if 'флагов' in full else 'CodeTable'
                    desc = ' '.join(sp['text'] for sp in line
                                    if 'Itali' in sp['font']).strip(' —–-\t')
                    cur = {'sec': s_n, 'num': t_n, 'type': ttype,
                           'title': desc, 'data': []}
                    cur_disc = cur_cat = None
                    cur_sub = ''
                    tables.append(cur)
                # skip ALL title-sized lines (including section headers)
                continue

            # ── italic title continuation (wraps from previous title line)
            if cur and all('Itali' in s['font'] and s['size'] >= TITLE_SZ
                           for s in line):
                cur['title'] += ' ' + ' '.join(s['text'] for s in line).strip()
                continue

            # ── discipline / category sub-header (SemiBold)
            if any(is_subhd(s) for s in line):
                subhdr_buf.append(full)
                continue

            # Flush buffered sub-header lines when non-subheader arrives
            if subhdr_buf:
                combined = ' '.join(subhdr_buf)
                subhdr_buf.clear()
                md = RE_DISC.search(combined)
                mc = RE_CAT.search(combined)
                if md:
                    new_disc = int(md.group(1))
                    if new_disc != cur_disc:
                        cur_cat = None   # reset cat when disc changes
                    cur_disc = new_disc
                if mc:
                    cur_cat = int(mc.group(1))
                cur_sub = combined

            # ── small-font lines (column headers, notes) → skip
            if all(is_small(s) for s in line):
                continue

            # ── data row (include superscripts ≥5pt for units exponents)
            if cur and any(is_data(s) for s in line):
                dspans = [s for s in line if s['size'] >= 5.0 and s['size'] < TITLE_SZ]
                if dspans:
                    cur['data'].append({
                        'spans': dspans,
                        'disc': cur_disc,
                        'cat': cur_cat,
                        'sub': cur_sub,
                    })

    doc.close()
    return tables


# ── Phase 2: Parse data spans into CSV rows ────────────────────────────

def parse_standard(tbl):
    """Parse standard 2-column CodeTable."""
    rows = []
    cur = None
    sec, num, title = tbl['sec'], tbl['num'], tbl['title']

    for d in tbl['data']:
        left  = [s for s in d['spans'] if s['x'] < STD_CODE_X]
        right = [s for s in d['spans'] if s['x'] >= STD_VAL_X]
        # Split notes column if present (x >= NOTES_X)
        note_sp = [s for s in right if s['x'] >= NOTES_X]
        if note_sp:
            right = [s for s in right if s['x'] < NOTES_X]

        lt = norm(' '.join(s['text'] for s in left).strip().replace(' ', ''))
        rt = ' '.join(s['text'] for s in right).strip()
        nt = ' '.join(s['text'] for s in note_sp).strip()

        if lt and is_num(lt):
            if cur:
                rows.append(cur)
            cur = _row(f'grib2/{sec}/{num}/{lt}', title, '', lt, '', rt, nt)
        elif cur:
            if rt:
                cur['MeaningParameterDescription_ru'] = append_text(
                    cur['MeaningParameterDescription_ru'], rt)
            if nt:
                cur['Note_ru'] = append_text(cur['Note_ru'], nt)

    if cur:
        rows.append(cur)
    return rows


def parse_flag(tbl):
    """Parse 3-column FlagTable (bit / value / meaning)."""
    rows = []
    cur = None
    prev_bit = ''
    sec, num, title = tbl['sec'], tbl['num'], tbl['title']

    for d in tbl['data']:
        bit_sp  = [s for s in d['spans'] if s['x'] < FLG_BIT_X]
        val_sp  = [s for s in d['spans'] if FLG_BIT_X <= s['x'] < FLG_MEAN_X]
        mean_sp = [s for s in d['spans'] if s['x'] >= FLG_MEAN_X]

        bt = norm(' '.join(s['text'] for s in bit_sp).strip().replace(' ', ''))
        vt = ' '.join(s['text'] for s in val_sp).strip()
        mt = ' '.join(s['text'] for s in mean_sp).strip()

        # If val-zone text isn't 0/1, it's meaning text (reserved ranges)
        if vt and vt not in ('0', '1'):
            mt = vt + (' ' + mt if mt else '')
            vt = ''

        if bt and is_num(bt):
            # New bit (with or without value)
            if cur:
                rows.append(cur)
            prev_bit = bt
            val = vt if vt in ('0', '1') else ''
            rid = f'grib2/{sec}/{num}/{bt}/{val}' if val else f'grib2/{sec}/{num}/{bt}'
            cur = _row(rid, title, '', bt, val, mt, '')

        elif vt in ('0', '1'):
            # Second value for same bit
            if cur:
                rows.append(cur)
            rid = f'grib2/{sec}/{num}/{prev_bit}/{vt}'
            cur = _row(rid, title, '', prev_bit, vt, mt, '')

        elif cur:
            # Multi-line continuation
            cont = mt if mt else ' '.join(s['text'] for s in d['spans']).strip()
            if cont:
                cur['MeaningParameterDescription_ru'] = append_text(
                    cur['MeaningParameterDescription_ru'], cont)

    if cur:
        rows.append(cur)
    return rows


def parse_table41(tbl):
    """Parse Table 4.1 — categories by discipline — single output file."""
    rows = []
    cur = None
    title = tbl['title']

    for d in tbl['data']:
        disc = d.get('disc')
        sub  = d.get('sub', '')
        left  = [s for s in d['spans'] if s['x'] < T4_DESC_X]
        right = [s for s in d['spans'] if s['x'] >= T4_DESC_X]

        lt = norm(' '.join(s['text'] for s in left).strip().replace(' ', ''))
        rt = ' '.join(s['text'] for s in right).strip()

        if lt and is_num(lt):
            if cur:
                rows.append(cur)
            rid = f'grib2/4/1/{disc}/{lt}' if disc is not None else f'grib2/4/1/{lt}'
            cur = _row(rid, title, sub, lt, '', rt, '')
        elif cur:
            if rt:
                cur['MeaningParameterDescription_ru'] = append_text(
                    cur['MeaningParameterDescription_ru'], rt)

    if cur:
        rows.append(cur)
    return rows


def parse_table42(tbl):
    """Parse Table 4.2 — parameters by discipline+category.
    Returns OrderedDict of (disc, cat) → [rows]."""
    groups = OrderedDict()
    cur = None
    title = tbl['title']

    for d in tbl['data']:
        disc = d.get('disc')
        cat  = d.get('cat')
        sub  = d.get('sub', '')
        left  = [s for s in d['spans'] if s['x'] < T4_DESC_X]
        mid   = [s for s in d['spans'] if T4_DESC_X <= s['x'] < T4_UNITS_X]
        right = [s for s in d['spans'] if s['x'] >= T4_UNITS_X]

        lt = norm(' '.join(s['text'] for s in left).strip().replace(' ', ''))
        mt = ' '.join(s['text'] for s in mid).strip()
        ut = ' '.join(s['text'] for s in right).strip()

        if lt and is_num(lt):
            if cur:
                key = (cur.pop('_d'), cur.pop('_c'))
                groups.setdefault(key, []).append(cur)
            rid = f'grib2/4/2/{disc}/{cat}/{lt}' if disc is not None and cat is not None else f'grib2/4/2/{lt}'
            cur = _row(rid, title, sub, lt, '', mt, '')
            cur['UnitComments_ru'] = ut
            cur['_d'] = disc
            cur['_c'] = cat
        elif cur:
            if mt:
                cur['MeaningParameterDescription_ru'] = append_text(
                    cur['MeaningParameterDescription_ru'], mt)
            if ut:
                cur['UnitComments_ru'] = append_text(cur['UnitComments_ru'], ut)

    if cur:
        key = (cur.pop('_d'), cur.pop('_c'))
        groups.setdefault(key, []).append(cur)

    return groups


def _row(rid, title, sub, cf, val, meaning, note):
    """Build a CSV row dict."""
    return {
        'ID': rid,
        'Title_ru': title,
        'SubTitle_ru': sub,
        'CodeFlag': cf,
        'Value': val,
        'MeaningParameterDescription_ru': meaning,
        'Note_ru': note,
        'noteIDs': '',
        'UnitComments_ru': '',
        'Status': '',
        'translated_by': 'Original',
    }


# ── Output ─────────────────────────────────────────────────────────────

def write_csv(fpath, rows):
    os.makedirs(os.path.dirname(fpath) if os.path.dirname(fpath) else '.', exist_ok=True)
    with open(fpath, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLS, extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)


def process_all(sections):
    os.makedirs(OUTDIR, exist_ok=True)
    total_f = total_r = 0

    for sec in sections:
        chunk_dir = os.path.join(CHUNKS, f'CodeTable_Sec{sec}')
        if not os.path.isdir(chunk_dir):
            print(f"  Sec {sec}: no chunk dir, skipping")
            continue

        print(f"\n=== Section {sec} ===")
        tables = walk_chunk(sec)
        print(f"  Tables detected: {len(tables)}")

        for tbl in tables:
            s, n, tt = tbl['sec'], tbl['num'], tbl['type']
            nd = len(tbl['data'])

            if tt == 'FlagTable':
                rows = parse_flag(tbl)
                fname = f'GRIB2_CodeFlag_{s}_{n}_FlagTable_ru.csv'
                write_csv(os.path.join(OUTDIR, fname), rows)
                print(f"    {fname:55s} {len(rows):4d} rows  (from {nd} data lines)")
                total_f += 1; total_r += len(rows)

            elif s == 4 and n == 1:
                rows = parse_table41(tbl)
                fname = 'GRIB2_CodeFlag_4_1_CodeTable_ru.csv'
                write_csv(os.path.join(OUTDIR, fname), rows)
                print(f"    {fname:55s} {len(rows):4d} rows  (from {nd} data lines)")
                total_f += 1; total_r += len(rows)

            elif s == 4 and n == 2:
                groups = parse_table42(tbl)
                for (disc, cat), grows in groups.items():
                    if not grows:
                        continue
                    fname = f'GRIB2_CodeFlag_4_2_{disc}_{cat}_CodeTable_ru.csv'
                    write_csv(os.path.join(OUTDIR, fname), grows)
                    print(f"    {fname:55s} {len(grows):4d} rows")
                    total_f += 1; total_r += len(grows)
                print(f"    Table 4.2 total: {sum(len(v) for v in groups.values())} rows in {len(groups)} files")

            else:
                rows = parse_standard(tbl)
                fname = f'GRIB2_CodeFlag_{s}_{n}_CodeTable_ru.csv'
                write_csv(os.path.join(OUTDIR, fname), rows)
                print(f"    {fname:55s} {len(rows):4d} rows  (from {nd} data lines)")
                total_f += 1; total_r += len(rows)

    print(f"\n{'='*60}")
    print(f"Total: {total_f} files, {total_r} rows → {OUTDIR}")


if __name__ == '__main__':
    secs = [int(x) for x in sys.argv[1:]] if len(sys.argv) > 1 else [0, 1, 3, 4, 5, 6]
    process_all(secs)
