#!/usr/bin/env python3
"""
Extract CCT (Common Code Tables) from Russian, French, Spanish PDFs.

Reads:  data/{lang}/chunks_cct/C{NN}.pdf
Writes: data/{lang}/cct_output/C{NN}_{lang}.csv

Each CCT table has a unique column schema. Column x-positions differ by language.

Usage:
    cd wmo_pipeline
    python tools/extract_cct.py ru          # one language
    python tools/extract_cct.py ru fr es    # all three
"""

import os, re, csv, sys
import pymupdf

HERE     = os.path.dirname(os.path.abspath(__file__))
PIPELINE = os.path.dirname(HERE)
EN_DIR   = '/Users/omard/Documents/projects/CCT-translation/english'

# ── Table schemas: (table_id, output_columns, lang_columns) ──────────
# lang_columns: column names that get _{lang} suffix
# All other columns are language-independent

TABLE_SCHEMAS = {
    'C00': {
        'cols': ['ID','GRIB version number','BUFR version number','CREX version number',
                 'Effective date','Status','translation_source'],
        'lang_cols': [],  # no translated columns — all are technical
    },
    'C01': {
        'cols': ['ID','CodeFigureForF1F2','CodeFigureForF3F3F3',
                 'Octet5GRIB1_Octet6BUFR3','OriginatingGeneratingCentres_{lang}',
                 'Status','translation_source'],
        'lang_cols': ['OriginatingGeneratingCentres'],
    },
    'C02': {
        'cols': ['ID','DateOfAssignment_{lang}','CodeFigureForrara','CodeFigureForBUFR',
                 'RadiosondeSoundingSystemUsed_{lang}','Status','translation_source'],
        'lang_cols': ['DateOfAssignment','RadiosondeSoundingSystemUsed'],
    },
    'C03': {
        'cols': ['ID','CodeFigureForIXIIXIX','CodeFigureForBUFR',
                 'InstrumentMakeAndType_{lang}','EquationCoefficients_a',
                 'EquationCoefficients_b','Status','translation_source'],
        'lang_cols': ['InstrumentMakeAndType'],
    },
    'C04': {
        'cols': ['ID','CodeFigureForXRXR','CodeFigureForBUFR',
                 'Meaning_{lang}','Status','translation_source'],
        'lang_cols': ['Meaning'],
    },
    'C05': {
        'cols': ['ID','CodeFigureForI6I6I6','CodeFigureForBUFR','CodeFigureForGRIB2',
                 'SatelliteName_{lang}','Status','translation_source'],
        'lang_cols': ['SatelliteName'],
    },
    'C06': {
        'cols': ['ID','CodeFigure','UnitType_{lang}','Meaning_{lang}','conventional',
                 'IA5-ASCII','ITA2','SIDefinition','Note','NoteID','Status','translation_source'],
        'lang_cols': ['UnitType','Meaning'],
    },
    'C07': {
        'cols': ['ID','CodeFigureForsasa','CodeFigureForBUFR',
                 'TrackingTechniquesStatusOfSystemUsed_{lang}','Status','translation_source'],
        'lang_cols': ['TrackingTechniquesStatusOfSystemUsed'],
    },
    'C08': {
        'cols': ['ID','Code','Agency_{lang}','Type_{lang}',
                 'InstrumentShortName_{lang}','InstrumentLongName_{lang}',
                 'Status','translation_source'],
        'lang_cols': ['Agency','Type','InstrumentShortName','InstrumentLongName'],
    },
    'C11': {
        'cols': ['ID','CREX2','GRIB2_BUFR4',
                 'OriginatingGeneratingCentre_{lang}','Status','translation_source'],
        'lang_cols': ['OriginatingGeneratingCentre'],
    },
    'C12': {
        'cols': ['ID','CodeFigure_OriginatingCentres','Name_OriginatingCentres_{lang}',
                 'CodeFigure_SubCentres','Name_SubCentres_{lang}','Status','translation_source'],
        'lang_cols': ['Name_OriginatingCentres','Name_SubCentres'],
    },
    'C13': {
        'cols': ['ID','CodeFigure_DataCategories','Name_DataCategories_{lang}',
                 'CodeFigure_InternationalDataSubcategories',
                 'Name_InternationalDataSubcategories_{lang}','Status','translation_source'],
        'lang_cols': ['Name_DataCategories','Name_InternationalDataSubcategories'],
    },
    'C14': {
        'cols': ['ID','CodeFigure','Meaning_{lang}','ChemicalFormula',
                 'Status','translation_source'],
        'lang_cols': ['Meaning'],
    },
}

# ── Column x-boundaries per language per table ───────────────────────
# Each entry: list of x-boundary values. Spans with x < boundaries[0] are col 0,
# spans with boundaries[0] <= x < boundaries[1] are col 1, etc.
# Number of boundaries = number of data columns - 1

COL_BOUNDS = {
    'C00': {'ru': [130, 225, 310, 420], 'fr': [130, 225, 310, 420], 'es': [130, 225, 310, 420]},
    'C01': {'ru': [135, 220, 310], 'fr': [150, 230, 325], 'es': [130, 205, 280]},
    'C02': {'ru': [115, 160, 250], 'fr': [115, 160, 250], 'es': [105, 150, 230]},
    'C03': {'ru': [115, 180, 380, 440], 'fr': [115, 180, 380, 440], 'es': [105, 170, 370, 430]},
    'C04': {'ru': [115, 180], 'fr': [115, 180], 'es': [105, 170]},
    'C05': {'ru': [135, 210, 270], 'fr': [135, 210, 270], 'es': [125, 200, 260]},
    'C06': {'ru': [110, 260, 325, 390, 450], 'fr': [120, 250, 330, 400, 450], 'es': [90, 225, 300, 390, 450]},
    'C07': {'ru': [115, 180], 'fr': [115, 180], 'es': [105, 170]},
    'C08': {'ru': [100, 155, 235, 340], 'fr': [108, 168, 255, 350], 'es': [80, 135, 243, 340]},
    'C11': {'ru': [145, 230], 'fr': [145, 230], 'es': [135, 220]},
    'C12': {'ru': [135, 280, 330], 'fr': [130, 275, 320], 'es': [100, 300, 340]},
    'C13': {'ru': [145, 280, 335], 'fr': [125, 275, 320], 'es': [100, 288, 315]},
    'C14': {'ru': [170, 410], 'fr': [150, 420], 'es': [120, 420]},
}

# ── Filter thresholds per language ───────────────────────────────────
FILTERS = {
    'ru': {'y_min': 90, 'y_max': 740, 'title_size': 9.5, 'data_size_lo': 7.0, 'data_size_hi': 10.5},
    'fr': {'y_min': 60, 'y_max': 710, 'title_size': 9.5, 'data_size_lo': 6.0, 'data_size_hi': 10.5},
    'es': {'y_min': 50, 'y_max': 790, 'title_size': 9.5, 'data_size_lo': 5.0, 'data_size_hi': 10.5},
}

# Continuation / header patterns to skip
SKIP_PATTERNS = re.compile(
    r'продолж\.|suite|a seguir|'
    r'ОБЩИЕ КОДОВЫЕ|TABLES DE CODE COMMUNES|TABLAS DE CIFRADO|'
    r'I\.2\s*[–—-]\s*(Общ|Co\d|Tab Co)',
    re.IGNORECASE
)

# ── Helpers ──────────────────────────────────────────────────────────

def norm_dash(s):
    return s.replace('\u2013', '-').replace('\u2014', '-')

def append_text(base, new):
    if not new: return base
    if not base: return new
    if base.endswith('-') and new and new[0].islower():
        return base[:-1] + new
    return base + ' ' + new

def get_spans(page, filt):
    out = []
    for blk in page.get_text("dict")["blocks"]:
        if "lines" not in blk: continue
        for line in blk["lines"]:
            for sp in line["spans"]:
                y = sp["bbox"][1]
                if y < filt['y_min'] or y > filt['y_max']:
                    continue
                t = sp["text"].strip()
                if not t: continue
                out.append({'x': sp["bbox"][0], 'y': sp["bbox"][1],
                            'text': t, 'font': sp["font"],
                            'size': round(sp["size"], 1)})
    return out

def group_y(spans, tol=3.0):
    if not spans: return []
    ss = sorted(spans, key=lambda s: (s['y'], s['x']))
    lines = [[ss[0]]]
    for s in ss[1:]:
        if abs(s['y'] - lines[-1][0]['y']) <= tol:
            lines[-1].append(s)
        else:
            lines.append([s])
    return [sorted(ln, key=lambda s: s['x']) for ln in lines]

def classify_spans(spans, bounds):
    """Classify spans into columns based on x-boundaries."""
    cols = [[] for _ in range(len(bounds) + 1)]
    for s in spans:
        x = s['x']
        placed = False
        for i, b in enumerate(bounds):
            if x < b:
                cols[i].append(s)
                placed = True
                break
        if not placed:
            cols[-1].append(s)
    return [' '.join(s['text'] for s in col) for col in cols]

def is_code_like(text):
    """Check if text looks like a code figure (number, range, or structured code)."""
    t = norm_dash(text.strip().replace(' ', ''))
    return bool(re.match(r'^\d[\d\-/]*$', t))


# ── Generic extraction with multi-line joining ───────────────────────

# Column header keywords (all languages) — lines matching these are skipped
COL_HEADER_KW = re.compile(
    r'кодовая цифра|номер бита|значение|наименование|дата назначения|'
    r'октет|форма ключа|системы радиозондирования|характеристик|'
    r'chiffre|numéro|signification|désignation|date|forme|'
    r'cifra|número|significado|designación|fecha|forma|'
    r'code figure|meaning|status|effective date|instrument|'
    r'спутник|agence|organisme|satellite|tracking',
    re.IGNORECASE
)

# Table title patterns — first 1-3 lines of first page
TITLE_KW = re.compile(
    r'общая кодовая таблица|table de code commune|tabla de cifrado comun|'
    r'common code table|C[–-]\d+',
    re.IGNORECASE
)

def extract_table(table_id, lang):
    """Extract one CCT table. Returns list of joined row dicts (col_idx → text)."""
    chunks_dir = os.path.join(PIPELINE, f'data/{lang}/chunks_cct')
    pdf_path = os.path.join(chunks_dir, f'{table_id}.pdf')
    if not os.path.exists(pdf_path):
        print(f"    SKIP {table_id}: no PDF")
        return []

    filt = FILTERS[lang]
    bounds = COL_BOUNDS[table_id][lang]
    n_cols = len(bounds) + 1
    doc = pymupdf.open(pdf_path)

    joined_rows = []  # list of lists, each = [col0, col1, ...]
    current = None
    data_started = False

    for pg in range(len(doc)):
        spans = get_spans(doc[pg], filt)
        lines = group_y(spans)

        for line in lines:
            full = ' '.join(s['text'] for s in line)

            # Skip continuation/header/footer patterns
            if SKIP_PATTERNS.search(full):
                continue

            # Skip table title lines (bold+large, or matching title keywords)
            if TITLE_KW.search(full):
                continue

            # Skip if ALL spans are bold (title/subtitle line)
            if all('Bold' in s['font'] for s in line) and not data_started:
                continue

            # Skip column header lines
            if COL_HEADER_KW.search(full) and not data_started:
                continue

            # Also skip column headers on continuation pages (re-detected by keyword match)
            if COL_HEADER_KW.search(full) and all(s['size'] < filt['title_size'] for s in line):
                continue

            # Classify spans into columns
            data_spans = [s for s in line if s['size'] >= filt['data_size_lo']]
            if not data_spans:
                continue

            col_vals = classify_spans(data_spans, bounds)

            # Detect new row vs continuation
            non_empty = [(j, v.strip()) for j, v in enumerate(col_vals) if v.strip()]
            has_bold = any('Bold' in s['font'] for s in data_spans)
            col0 = norm_dash(col_vals[0].strip()) if col_vals else ''

            # New row if: multi-column, or col0 has text, or bold group header
            if len(non_empty) >= 2:
                is_new = True
            elif len(non_empty) == 1 and non_empty[0][0] == 0:
                is_new = True  # code-only row
            elif len(non_empty) == 1 and has_bold and data_started:
                is_new = True  # bold group header
            else:
                is_new = False  # single-column regular text → continuation

            if is_new:
                if current:
                    joined_rows.append(current)
                current = col_vals
                data_started = True
            elif current and non_empty:
                # Continuation — append text to each non-empty column
                for j in range(n_cols):
                    if j < len(col_vals) and col_vals[j].strip():
                        current[j] = append_text(current[j], col_vals[j].strip())
            # else: skip (no current row yet, still in headers)

    if current:
        joined_rows.append(current)

    doc.close()
    return joined_rows


def build_csv_rows(table_id, lang, raw_rows):
    """Match extracted rows to English reference and build output CSV rows."""
    schema = TABLE_SCHEMAS[table_id]
    cols = [c.replace('{lang}', lang) for c in schema['cols']]

    # Read English reference
    en_path = os.path.join(EN_DIR, f'{table_id}_en.csv')
    en_rows = []
    en_cols = []
    if os.path.exists(en_path):
        with open(en_path) as f:
            reader = csv.DictReader(f)
            en_cols = reader.fieldnames
            en_rows = list(reader)

    csv_rows = []

    # Match by position (row index) — works because code figure order matches English
    for i, raw in enumerate(raw_rows):
        row = {c: '' for c in cols}
        row['translation_source'] = 'PDF'

        # Copy structure from English if available
        if i < len(en_rows):
            en = en_rows[i]
            row['ID'] = en.get('ID', '')
            row['Status'] = en.get('Status', 'Operational')
            # Copy ALL language-independent columns from English
            for ec in en_cols:
                if ec in cols and '_{' not in ec and ec not in ('ID', 'Status', 'translation_source'):
                    row[ec] = en.get(ec, '')
                # Also copy non-lang columns that exist in output
                if ec in row and ec not in ('ID', 'Status', 'translation_source'):
                    if not any(ec.endswith(f'_{lang}') or ec.endswith('_{lang}') for _ in [1]):
                        if ec in en:
                            row[ec] = en[ec]

        # Fill extracted columns — skip column 0 (code figure, use English ID instead)
        # Map: raw[0]=code, raw[1]=first content col, raw[2]=second content col, ...
        # cols: [ID, ...structural..., ...lang_cols..., Status, translation_source]
        # The structural/lang column order matches the raw column order
        content_cols = [c for c in cols if c not in ('ID', 'Status', 'translation_source')]
        for j, val in enumerate(raw):
            cleaned = norm_dash(val.strip()) if val else ''
            if j < len(content_cols):
                col_name = content_cols[j]
                row[col_name] = cleaned

        csv_rows.append(row)

    return csv_rows, cols


# ── Main ─────────────────────────────────────────────────────────────

def process_language(lang):
    out_dir = os.path.join(PIPELINE, f'data/{lang}/cct_output')
    os.makedirs(out_dir, exist_ok=True)

    total_files = 0
    total_rows = 0
    tables = ['C00','C01','C02','C03','C04','C05','C06','C07','C08','C11','C12','C13','C14']

    for tbl in tables:
        raw = extract_table(tbl, lang)
        if not raw:
            continue

        csv_rows, cols = build_csv_rows(tbl, lang, raw)

        out_path = os.path.join(out_dir, f'{tbl}_{lang}.csv')
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=cols, lineterminator='\r\n')
            w.writeheader()
            w.writerows(csv_rows)

        print(f"    {tbl}_{lang}.csv: {len(csv_rows):4d} rows (extracted {len(raw)} lines)")
        total_files += 1
        total_rows += len(csv_rows)

    print(f"  {lang.upper()} total: {total_files} files, {total_rows} rows")
    return total_files, total_rows


if __name__ == '__main__':
    langs = sys.argv[1:] if len(sys.argv) > 1 else ['ru', 'fr', 'es']
    grand_files = grand_rows = 0
    for lang in langs:
        print(f"\n=== {lang.upper()} ===")
        f, r = process_language(lang)
        grand_files += f
        grand_rows += r
    print(f"\n{'='*50}")
    print(f"Grand total: {grand_files} files, {grand_rows} rows")
