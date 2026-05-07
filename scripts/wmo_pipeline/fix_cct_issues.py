#!/usr/bin/env python3
"""Fix 3 high-priority CCT issues:
1. C12_fr header contamination (95 rows)
2. C08 tab/column bleed (36 cells across RU and FR)
3. C08 truncated words (Russian, ~60 cells ending with hyphen)
"""

import csv
import re
from pathlib import Path

# Paths
C12_FR = Path("/Users/omard/Documents/projects/CCT-translation/french/C12_fr.csv")
C08_RU = Path("/Users/omard/Documents/projects/CCT-translation/russian/C08_ru.csv")
C08_FR = Path("/Users/omard/Documents/projects/CCT-translation/french/C08_fr.csv")
C08_EN = Path("/Users/omard/Documents/projects/CCT-translation/english/C08_en.csv")
C08_PDF = Path("/Users/omard/Documents/projects/WMO_work_claude/table_extractor_vision/wmo_pipeline/data/ru/chunks_cct/C08.pdf")

# Known compound-word continuations in this dataset.
# These are complete Russian words that appear after a compound-word hyphen
# (e.g., "радиолокатор-высотомер"). All other lowercase continuations
# after a hyphen are line-break word splits.
COMPOUND_WORDS = {
    "высотомер",   # altimeter (in "радиолокатор-высотомер")
    # Add more as needed
}


def read_csv(path):
    """Read CSV preserving all fields."""
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        header = next(reader)
        rows = list(reader)
    return header, rows


def write_csv(path, header, rows):
    """Write CSV."""
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


# ============================================================
# FIX 1: C12_fr header contamination
# ============================================================
def fix1_c12_header_contamination():
    """Strip page header text from Name_OriginatingCentres_fr column."""
    header, rows = read_csv(C12_FR)
    name_col = header.index("Name_OriginatingCentres_fr")

    HEADER_PATTERN = re.compile(
        r'\s*C\s+entres\s+d\s*\'\s*origine\s+Octet.*?(?:édition\s+\d+\))',
        re.DOTALL
    )

    fix_count = 0
    for row in rows:
        val = row[name_col]
        if 'entres d' not in val:
            continue

        match = HEADER_PATTERN.search(val)
        if not match:
            continue

        clean_name = val[:match.start()].strip()
        if clean_name != val:
            row[name_col] = clean_name
            fix_count += 1

    write_csv(C12_FR, header, rows)
    print(f"FIX 1: Fixed {fix_count} rows in C12_fr.csv (header contamination)")
    return fix_count


# ============================================================
# FIX 2: C08 tab/column bleed
# ============================================================
def fix2_c08_tab_bleed():
    """Fix tab characters causing column bleed in C08 RU and FR."""
    total_fixes = 0

    for label, csv_path in [("C08_ru", C08_RU), ("C08_fr", C08_FR)]:
        header, rows = read_csv(csv_path)
        file_fixes = 0

        text_cols = [i for i, h in enumerate(header)
                     if h not in ('ID', 'Code', 'Status', 'translation_source')]

        for row in rows:
            for col_idx in text_cols:
                val = row[col_idx]
                if '\t' not in val:
                    continue

                parts = val.split('\t')
                first_part = parts[0].rstrip()
                second_part = '\t'.join(parts[1:]).lstrip()

                row[col_idx] = first_part

                next_col_idx = None
                for tc in text_cols:
                    if tc > col_idx:
                        next_col_idx = tc
                        break

                if next_col_idx is not None:
                    if not row[next_col_idx]:
                        row[next_col_idx] = second_part
                    else:
                        row[next_col_idx] = second_part + ' ' + row[next_col_idx]

                file_fixes += 1

        # Fix C08/776 RU: "РРОСКОСМОС" -> "РОСКОСМОС"
        if label == "C08_ru":
            agency_col = header.index("Agency_ru")
            for row in rows:
                if row[agency_col] == "РРОСКОСМОС":
                    row[agency_col] = "РОСКОСМОС"
                    file_fixes += 1
                    print(f"  Fixed РРОСКОСМОС -> РОСКОСМОС typo")

        # Fix C08/606 FR: longname has merged 607 data
        if label == "C08_fr":
            longname_col = header.index("InstrumentLongName_fr")
            for row in rows:
                if row[0] == "C08/606":
                    ln = row[longname_col]
                    if "607 NOAA" in ln:
                        row[longname_col] = "Sondeur infrarouge perfectionné haute résolution/3"
                        file_fixes += 1
                        print(f"  Fixed C08/606 InstrumentLongName_fr")

        write_csv(csv_path, header, rows)
        print(f"FIX 2: Fixed {file_fixes} cells in {label}.csv (tab/column bleed)")
        total_fixes += file_fixes

    return total_fixes


# ============================================================
# FIX 3: C08 truncated words (Russian)
# ============================================================
def _is_line_break_hyphen(before_text, after_text):
    """Determine if a hyphen between before_text and after_text is a line break.

    Returns True if the hyphen should be removed (it's a line break),
    False if it should be kept (it's a compound word connector).

    Strategy: if the continuation starts with a lowercase letter, it's a line
    break UNLESS the first word is a known compound-word component.
    """
    if not after_text:
        return False

    # Strip trailing punctuation for word check
    first_word = after_text.split()[0] if after_text.split() else after_text
    clean_word = first_word.rstrip('.,;:!?)')

    # If continuation starts with uppercase, NOT a line break
    if first_word and first_word[0].isupper():
        return False

    # If continuation starts with lowercase, it's a line break
    # UNLESS it's a known compound word
    if first_word and first_word[0].islower():
        if clean_word in COMPOUND_WORDS:
            return False
        return True

    return False


def _extract_pdf_table(pdf_path):
    """Extract full text per code per column from C08 PDF.

    Returns dict: code_str -> {col_name: full_text}
    """
    import pymupdf

    doc = pymupdf.open(str(pdf_path))

    # Collect ALL spans across all pages, excluding footers
    FOOTER_PATTERNS = re.compile(r'(I\.2\s*–\s*Общ\.|продолж\.|^\($|^\)$)')

    all_spans = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if "lines" not in block:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if not text:
                        continue
                    bbox = span["bbox"]
                    y0 = bbox[1]

                    # Skip footer area (y > 740)
                    if y0 > 740:
                        continue

                    # Skip continuation markers and footers
                    if FOOTER_PATTERNS.search(text):
                        continue

                    # Strip tabs from spans (tab = column bleed in PDF)
                    if '\t' in text:
                        text = text.split('\t')[0].rstrip()
                        if not text:
                            continue

                    all_spans.append({
                        "page": page_num,
                        "text": text,
                        "x0": bbox[0],
                        "y0": y0,
                    })
    doc.close()

    # Sort by page, y, x
    all_spans.sort(key=lambda s: (s["page"], s["y0"], s["x0"]))

    def classify_col(x0):
        if x0 < 100:
            return "Code"
        elif x0 < 155:
            return "Agency"
        elif x0 < 240:
            return "Type"
        elif x0 < 340:
            return "ShortName"
        else:
            return "LongName"

    # Find all NUMERIC code spans (x < 100, numeric text)
    code_spans = []
    for i, span in enumerate(all_spans):
        if classify_col(span["x0"]) == "Code" and span["text"].isdigit():
            code_spans.append((i, span))

    # For each code, collect spans between this code and the next
    result = {}
    for ci, (span_idx, code_span) in enumerate(code_spans):
        code = code_span["text"]
        code_page = code_span["page"]
        code_y = code_span["y0"]

        # Find y-boundary
        if ci + 1 < len(code_spans):
            next_cs = code_spans[ci + 1][1]
            if next_cs["page"] == code_page:
                max_y = next_cs["y0"] - 1
            else:
                max_y = 9999
        else:
            max_y = 9999

        # Collect spans belonging to this code
        col_texts = {"Agency": [], "Type": [], "ShortName": [], "LongName": []}
        for span in all_spans:
            if span["page"] != code_page:
                continue
            if span["y0"] < code_y - 1:
                continue
            if span["y0"] > max_y:
                continue
            col = classify_col(span["x0"])
            if col in col_texts:
                col_texts[col].append(span["text"])

        # Join multi-line text with smart hyphen handling
        def join_lines(texts):
            if not texts:
                return ""
            result_str = texts[0]
            for t in texts[1:]:
                if result_str.endswith('-'):
                    if _is_line_break_hyphen(result_str, t):
                        # Line break: remove hyphen, concatenate directly
                        result_str = result_str[:-1] + t
                    else:
                        # Compound word: keep hyphen, concatenate directly
                        result_str = result_str + t
                else:
                    result_str = result_str + ' ' + t

            # Clean double spaces
            result_str = re.sub(r'\s+', ' ', result_str).strip()
            return result_str

        entry = {}
        for col_name, texts in col_texts.items():
            entry[col_name] = join_lines(texts)

        result[code] = entry

    return result


def fix3_c08_truncated_words():
    """Fix truncated Russian words by reading full text from PDF."""
    header, rows = read_csv(C08_RU)

    csv_to_pdf_col = {
        "Agency_ru": "Agency",
        "Type_ru": "Type",
        "InstrumentShortName_ru": "ShortName",
        "InstrumentLongName_ru": "LongName",
    }

    text_cols = {h: header.index(h) for h in csv_to_pdf_col}

    # Truncated = Cyrillic letter followed by hyphen at end
    TRUNCATED = re.compile(r'[а-яА-ЯёЁ]-$')

    truncated_cells = []
    for row_idx, row in enumerate(rows):
        for col_name, col_idx in text_cols.items():
            val = row[col_idx].strip()
            if val and TRUNCATED.search(val):
                truncated_cells.append((row_idx, col_name, col_idx, val))

    print(f"\nFIX 3: Found {len(truncated_cells)} truncated cells in C08_ru.csv")

    # Extract full text from PDF
    pdf_data = _extract_pdf_table(C08_PDF)

    fix_count = 0
    unfixed = []

    for row_idx, col_name, col_idx, val in truncated_cells:
        row = rows[row_idx]
        csv_id = row[0]
        code_str = row[1]
        pdf_col = csv_to_pdf_col[col_name]

        if code_str not in pdf_data:
            unfixed.append((csv_id, col_name, val, "code not found in PDF"))
            continue

        pdf_text = pdf_data[code_str].get(pdf_col, "")
        if not pdf_text:
            unfixed.append((csv_id, col_name, val, "column empty in PDF"))
            continue

        # Verify the truncated text is a prefix of the PDF text
        truncated_stem = val[:-1]  # Remove trailing hyphen

        def normalize(s):
            return re.sub(r'\s+', ' ', s).strip()

        pdf_norm = normalize(pdf_text)
        stem_norm = normalize(truncated_stem)

        # Handle tab-contaminated stems (from tabs that were in original CSV)
        clean_stem = stem_norm.split('\t')[0].strip()

        if pdf_norm.startswith(clean_stem) or pdf_norm.startswith(stem_norm):
            new_val = pdf_text
            rows[row_idx][col_idx] = new_val
            fix_count += 1
            print(f"  {csv_id} [{col_name}]: '{val}' -> '{new_val}'")
        else:
            unfixed.append((csv_id, col_name, val,
                           f"stem mismatch: '{clean_stem[:30]}' vs '{pdf_norm[:50]}'"))

    if unfixed:
        print(f"\n  Could not fix {len(unfixed)} cells:")
        for csv_id, col_name, val, reason in unfixed:
            print(f"    {csv_id} [{col_name}]: '{val}' -- {reason}")

    write_csv(C08_RU, header, rows)
    print(f"\nFIX 3: Fixed {fix_count}/{len(truncated_cells)} truncated cells in C08_ru.csv")
    return fix_count, len(truncated_cells) - fix_count


# ============================================================
# Verification
# ============================================================
def verify_no_new_issues():
    """Check that fixes didn't introduce new problems."""
    issues = []

    # Check C12_fr: no more header contamination
    header, rows = read_csv(C12_FR)
    name_col = header.index("Name_OriginatingCentres_fr")
    contam = sum(1 for r in rows if 'entres d' in r[name_col])
    if contam:
        issues.append(f"C12_fr still has {contam} contaminated rows")

    # Check C08_ru: no more tabs
    header, rows = read_csv(C08_RU)
    text_cols = [i for i, h in enumerate(header)
                 if h not in ('ID', 'Code', 'Status', 'translation_source')]
    tab_rows = []
    for r in rows:
        for c in text_cols:
            if '\t' in r[c]:
                tab_rows.append(f"  {r[0]} col {header[c]}: '{r[c][:60]}...'")
    if tab_rows:
        issues.append(f"C08_ru still has {len(tab_rows)} tab characters:")
        issues.extend(tab_rows)

    # Check C08_fr: no more tabs
    header, rows = read_csv(C08_FR)
    text_cols = [i for i, h in enumerate(header)
                 if h not in ('ID', 'Code', 'Status', 'translation_source')]
    tabs = sum(1 for r in rows for c in text_cols if '\t' in r[c])
    if tabs:
        issues.append(f"C08_fr still has {tabs} tab characters")

    # Check C08_ru: remaining truncated cells
    TRUNCATED = re.compile(r'[а-яА-ЯёЁ]-$')
    header, rows = read_csv(C08_RU)
    text_cols = [i for i, h in enumerate(header)
                 if h not in ('ID', 'Code', 'Status', 'translation_source')]
    trunc_rows = []
    for r in rows:
        for c in text_cols:
            val = r[c].strip()
            if val and TRUNCATED.search(val):
                trunc_rows.append(f"  {r[0]} col {header[c]}: '{val}'")
    if trunc_rows:
        issues.append(f"C08_ru still has {len(trunc_rows)} truncated cells:")
        issues.extend(trunc_rows)

    # Check C08_ru: РРОСКОСМОС typo
    header, rows = read_csv(C08_RU)
    agency_col = header.index("Agency_ru")
    typo = sum(1 for r in rows if "РРОСКОСМОС" in r[agency_col])
    if typo:
        issues.append(f"C08_ru still has {typo} РРОСКОСМОС typos")

    return issues


# ============================================================
# Main
# ============================================================
if __name__ == "__main__":
    print("=" * 60)
    print("CCT Issue Fixes")
    print("=" * 60)

    print("\n--- Fix 1: C12_fr header contamination ---")
    n1 = fix1_c12_header_contamination()

    # Fix 2 runs BEFORE Fix 3: splits tabs in CSV
    # Fix 3 then replaces truncated text with PDF text (tabs stripped from PDF)
    print("\n--- Fix 2: C08 tab/column bleed ---")
    n2 = fix2_c08_tab_bleed()

    print("\n--- Fix 3: C08 truncated words ---")
    n3, n3_unfixed = fix3_c08_truncated_words()

    print("\n--- Verification ---")
    issues = verify_no_new_issues()
    if issues:
        print("ISSUES FOUND:")
        for iss in issues:
            print(f"  {iss}")
    else:
        print("All clean - no residual issues found.")

    print("\n" + "=" * 60)
    print(f"SUMMARY: {n1 + n2 + n3} total fixes applied")
    print(f"  Fix 1: {n1} cells (C12_fr header contamination)")
    print(f"  Fix 2: {n2} cells (C08 tab/column bleed)")
    print(f"  Fix 3: {n3} cells fixed, {n3_unfixed} unfixed (C08 truncated words)")
    print("=" * 60)
