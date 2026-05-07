"""Extract BUFR Table D — Common sequences (FXY1/FXY2 groupings).

Table D is structurally different from Tables A/B/C:
- Each row has TWO FXY codes: FXY1 (sequence descriptor, starts with 3)
  and FXY2 (component descriptor — can be any F).
- One FXY1 maps to N FXY2 entries (its components).
- FXY1 sequences can span multiple pages; continuation is tracked via
  ``previous_fxy1``.
- The Id format includes an occurrence counter (1-based per FXY1 group):
  ``bufr4/3/{XX}/{YYY}/{occ}/{fxy2:06d}``

FXY span formats encountered across PDFs
-----------------------------------------
1. Full "F XX YYY" in one span  → most common
2. "F XX" + "YYY" in two spans at same y
3. "F" + "XX" + "YYY" in three spans at same y
4. "F" + "XX YYY" in two spans at same y
5. Standalone "XX YYY" (Spanish PDF omits F for continuation FXY1 entries)

Public functions
----------------
  extract_table_d(pdf_path, start_page, end_page, lang, lang_config)
      → dict[category_no, pd.DataFrame]
  save_table_d(results, output_dir, lang) → int
"""

from __future__ import annotations

import re
from typing import Optional
import fitz
import pandas as pd

from lang_config import LangConfig


def _lang_col(base: str, lang: str) -> str:
    return f"{base}_{lang}"


def _clean_text(text: str) -> str:
    """Normalize PDF extraction artifacts for Table D text fields."""
    text = text.replace("\x08", "").replace("\x07", "")
    text = text.replace("\t", " ")
    text = re.sub(r"(\w)- (\w)", r"\1-\2", text)
    text = re.sub(r" (–\d)", r"\1", text)
    text = text.replace("\u00ad ", "").replace("\u00ad", "")
    text = re.sub(r"  +", " ", text)
    # Strip footer remnants (all supported languages)
    text = re.sub(r"\s*\(à suivre\)\s*I\.2\s*–.*$", "", text)
    text = re.sub(r"\s*\(continued\)\s*I\.2\s*–.*$", "", text)
    text = re.sub(r"\s*\(continued\)\s*$", "", text)
    text = re.sub(r"\s*\(продолж\.\)\s*I\.2\s*–.*$", "", text)
    text = re.sub(r"\s*\(continúa\)\s*I\.2\s*–.*$", "", text)
    text = re.sub(r"\s*\(continúa\)\s*$", "", text)
    # Standalone page footer references
    text = re.sub(
        r"\s*I\.2\s*[–-]\s*(?:BUFR\s+)?Tabl[eao]\s+(?:BUFR\s+)?[A-D]\s*/?\d*\s*[–—-]\s*\d+\s*$",
        "", text,
    )
    return text.strip()


def _get_spans(page) -> list[dict]:
    """Extract all text spans with positions from a page."""
    blocks = page.get_text("dict")
    spans = []
    for block in blocks["blocks"]:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if text:
                    spans.append({
                        "text": text,
                        "x0": span["bbox"][0],
                        "y0": span["bbox"][1],
                    })
    return spans


def _classify_column(x0: float, col_bounds: dict) -> str | None:
    for col_name, (lo, hi) in col_bounds.items():
        if lo <= x0 < hi:
            return col_name
    return None


# ── FXY span parsing ──────────────────────────────────────────────────────────
# Table D FXY codes appear in many split/merged variants across PDFs.

FXY_FULL = re.compile(r"^(\d)\s+(\d{2})\s+(\d{3})$")
FXY_PARTIAL_FX = re.compile(r"^(\d)\s+(\d{2})$")
FXY_PARTIAL_Y = re.compile(r"^(\d{3})$")
FXY_SINGLE_F = re.compile(r"^(\d)$")
FXY_SINGLE_XX = re.compile(r"^(\d{2})$")
FXY_PARTIAL_XX_YYY = re.compile(r"^(\d{2})\s+(\d{3})$")


def _parse_fxy_text(text: str) -> str | None:
    """Parse 'F XX YYY' into 6-digit FXY string, or None."""
    m = FXY_FULL.match(text)
    if m:
        return f"{int(m.group(1)):01d}{int(m.group(2)):02d}{int(m.group(3)):03d}"
    return None


def _merge_fxy_spans(spans: list[dict], default_f: str | None = None) -> list[dict]:
    """Merge split FXY spans into dicts {fxy: str, y: float}.

    Handles five span formats:
    1. Full "F XX YYY" in one span
    2. "F XX" + "YYY" in two spans at same y
    3. "F" + "XX" + "YYY" in three spans at same y
    4. "F" + "XX YYY" in two spans at same y
    5. Standalone "XX YYY" with *default_f* (Spanish PDF omits F for
       continuation FXY1 entries on the same page)
    """
    entries = []
    spans_sorted = sorted(spans, key=lambda s: (s["y0"], s["x0"]))

    i = 0
    while i < len(spans_sorted):
        s = spans_sorted[i]
        fxy = _parse_fxy_text(s["text"])
        if fxy:
            entries.append({"fxy": fxy, "y": s["y0"]})
            i += 1
            continue

        m_fx = FXY_PARTIAL_FX.match(s["text"])
        if m_fx and i + 1 < len(spans_sorted):
            next_s = spans_sorted[i + 1]
            m_y = FXY_PARTIAL_Y.match(next_s["text"])
            if m_y and abs(next_s["y0"] - s["y0"]) < 3:
                fxy = f"{int(m_fx.group(1)):01d}{int(m_fx.group(2)):02d}{int(m_y.group(1)):03d}"
                entries.append({"fxy": fxy, "y": s["y0"]})
                i += 2
                continue

        m_f = FXY_SINGLE_F.match(s["text"])
        if m_f:
            # "F" + "XX YYY"
            if i + 1 < len(spans_sorted):
                s2 = spans_sorted[i + 1]
                m_xx_yyy = FXY_PARTIAL_XX_YYY.match(s2["text"])
                if m_xx_yyy and abs(s2["y0"] - s["y0"]) < 3:
                    fxy = f"{int(m_f.group(1)):01d}{int(m_xx_yyy.group(1)):02d}{int(m_xx_yyy.group(2)):03d}"
                    entries.append({"fxy": fxy, "y": s["y0"]})
                    i += 2
                    continue
            # "F" + "XX" + "YYY"
            if i + 2 < len(spans_sorted):
                s2, s3 = spans_sorted[i + 1], spans_sorted[i + 2]
                m_xx = FXY_SINGLE_XX.match(s2["text"])
                m_yyy = FXY_PARTIAL_Y.match(s3["text"])
                if (m_xx and m_yyy
                        and abs(s2["y0"] - s["y0"]) < 3
                        and abs(s3["y0"] - s["y0"]) < 3):
                    fxy = f"{int(m_f.group(1)):01d}{int(m_xx.group(1)):02d}{int(m_yyy.group(1)):03d}"
                    entries.append({"fxy": fxy, "y": s["y0"]})
                    i += 3
                    continue

        # Standalone "XX YYY" with default F (Spanish continuation entries)
        if default_f is not None:
            m_xx_yyy = FXY_PARTIAL_XX_YYY.match(s["text"])
            if m_xx_yyy:
                fxy = f"{int(default_f):01d}{int(m_xx_yyy.group(1)):02d}{int(m_xx_yyy.group(2)):03d}"
                entries.append({"fxy": fxy, "y": s["y0"]})
                i += 1
                continue

        i += 1

    return entries


# ── Page-level extraction ─────────────────────────────────────────────────────

def _extract_table_d_page(page, lang: str, col_bounds: dict,
                           previous_fxy1: str | None) -> tuple[list[dict], str | None]:
    """Extract Table D rows from one PDF page.

    Returns (flat_rows, last_fxy1) where each row has:
    FXY1, FXY2, ElementName_{lang}, ElementDescription_{lang}, Title_{lang}
    """
    spans = _get_spans(page)

    col_spans: dict[str, list] = {col: [] for col in col_bounds}
    for s in spans:
        col = _classify_column(s["x0"], col_bounds)
        if col:
            col_spans[col].append(s)
    for col in col_spans:
        col_spans[col].sort(key=lambda s: s["y0"])

    # FXY1 entries must start with "3"; pass default_f="3" for Spanish continuation
    fxy1_entries = [e for e in _merge_fxy_spans(col_spans["fxy1"], default_f="3")
                    if e["fxy"].startswith("3")]
    fxy2_entries = _merge_fxy_spans(col_spans["fxy2"])

    if not fxy2_entries:
        return [], previous_fxy1

    if previous_fxy1:
        if not fxy1_entries:
            fxy1_entries = [{"fxy": previous_fxy1, "y": 0}]
        else:
            first_y = fxy1_entries[0]["y"]
            if any(e["y"] < first_y - 1 for e in fxy2_entries):
                fxy1_entries = [{"fxy": previous_fxy1, "y": 0}] + fxy1_entries

    page_bottom = page.rect.height

    def _collect(col_name: str, y_start: float, y_end: float) -> str:
        texts = [s["text"] for s in col_spans[col_name]
                 if y_start - 1 <= s["y0"] <= y_end]
        return " ".join(texts)

    # ── Extract parenthesized titles from the gap above each FXY1 ──
    fxy1_titles: dict[str, str] = {}
    title_y_starts: dict[str, float] = {}

    for i, f1 in enumerate(fxy1_entries):
        y_above = fxy1_entries[i - 1]["y"] + 2 if i > 0 else 0
        y_below = f1["y"] - 1
        gap_spans = sorted(
            [s for s in col_spans["name"] if y_above <= s["y0"] <= y_below],
            key=lambda s: s["y0"],
        )
        title_spans = []
        in_title = False
        for s in gap_spans:
            if not in_title and s["text"].startswith("("):
                in_title = True
            if in_title:
                title_spans.append(s)
                if s["text"].endswith(")"):
                    break
        if title_spans:
            title_text = _clean_text(" ".join(s["text"] for s in title_spans))
            fxy1_titles[f1["fxy"]] = title_text
            title_y_starts[f1["fxy"]] = title_spans[0]["y0"] - 1
        else:
            fxy1_titles[f1["fxy"]] = ""

    # ── Assign y-ranges for FXY2, capped at section boundaries ──
    boundary_ys = sorted(
        list(title_y_starts.values()) + [f1["y"] - 3 for f1 in fxy1_entries]
    )
    for i, entry in enumerate(fxy2_entries):
        raw_y_end = fxy2_entries[i + 1]["y"] - 1 if i + 1 < len(fxy2_entries) else page_bottom
        cap = raw_y_end
        for b in boundary_ys:
            if entry["y"] + 2 < b <= raw_y_end:
                cap = min(cap, b)
                break
        entry["y_end"] = cap

    # ── Build rows ──
    rows = []
    current_fxy1 = None
    for entry in fxy2_entries:
        for f1 in reversed(fxy1_entries):
            if f1["y"] <= entry["y"] + 1:
                current_fxy1 = f1["fxy"]
                break
        if current_fxy1 is None:
            continue

        name = _clean_text(_collect("name", entry["y"], entry["y_end"]))
        description = _clean_text(_collect("description", entry["y"], entry["y_end"]))
        name = re.sub(r"(\w)- (\w)", r"\1\2", name)
        description = re.sub(r"(\w)- (\w)", r"\1\2", description)
        title = fxy1_titles.get(current_fxy1, "")

        rows.append({
            "FXY1": current_fxy1,
            "FXY2": entry["fxy"],
            _lang_col("ElementName", lang): name,
            _lang_col("ElementDescription", lang): description,
            _lang_col("Title", lang): title,
        })

    last_fxy1 = fxy1_entries[-1]["fxy"] if fxy1_entries else previous_fxy1
    return rows, last_fxy1


# ── Category detection ────────────────────────────────────────────────────────

def _detect_categories(
    pdf_path: str, start_page: int, end_page: int,
    category_pattern: re.Pattern, continuation_filter: str,
) -> list[dict]:
    """Detect Table D category boundaries from page text."""
    doc = fitz.open(pdf_path)
    raw = []
    for page_num in range(start_page, end_page + 1):
        text = doc[page_num].get_text()
        for m in category_pattern.finditer(text):
            name = m.group(2).strip()
            if continuation_filter not in name:
                raw.append({"category": m.group(1).zfill(2), "name": name, "page": page_num})
    doc.close()

    seen: set[str] = set()
    categories = []
    for c in raw:
        if c["category"] not in seen:
            seen.add(c["category"])
            categories.append(c)

    for i, c in enumerate(categories):
        c["start_page"] = c["page"]
        c["end_page"] = categories[i + 1]["page"] - 1 if i + 1 < len(categories) else end_page
    return categories


# ── Public extraction function ────────────────────────────────────────────────

def extract_table_d(
    pdf_path: str,
    start_page: int,
    end_page: int,
    lang: str,
    lang_config: LangConfig,
) -> dict[str, pd.DataFrame]:
    """Extract all Table D categories from PDF pages (0-indexed).

    Returns dict mapping category_no → DataFrame with columns:
    Id, Category, CategoryOfSequences_{lang}, FXY1, Title_{lang},
    SubTitle_{lang}, FXY2, ElementName_{lang}, ElementDescription_{lang},
    Note_{lang}, noteIDs, Status

    The ``Id`` format is ``bufr4/3/{XX}/{YYY}/{occ}/{FXY2:06d}`` where
    ``occ`` is the 1-based occurrence index within the FXY1 group.
    """
    categories = _detect_categories(
        pdf_path, start_page, end_page,
        lang_config.category_pattern, lang_config.continuation_filter,
    )
    doc = fitz.open(pdf_path)
    results: dict[str, pd.DataFrame] = {}

    for cat in categories:
        all_rows: list[dict] = []
        previous_fxy1 = None
        for page_num in range(cat["start_page"], cat["end_page"] + 1):
            rows, previous_fxy1 = _extract_table_d_page(
                doc[page_num], lang, lang_config.col_bounds_d, previous_fxy1
            )
            all_rows.extend(rows)

        if not all_rows:
            continue

        df = pd.DataFrame(all_rows)
        df = df[df["FXY1"].str[1:3] == cat["category"]].copy()
        if df.empty:
            continue

        cat_col = _lang_col("CategoryOfSequences", lang)
        title_col = _lang_col("Title", lang)
        subtitle_col = _lang_col("SubTitle", lang)
        name_col = _lang_col("ElementName", lang)
        desc_col = _lang_col("ElementDescription", lang)
        note_col = _lang_col("Note", lang)

        df["Category"] = cat["category"]
        if lang_config.category_name_prefix:
            df[cat_col] = f"{lang_config.category_name_prefix} {cat['category']} \u2014 {cat['name']}"
        else:
            df[cat_col] = cat["name"]
        df[subtitle_col] = ""
        df[note_col] = ""
        df["noteIDs"] = ""
        df["Status"] = "Operational"

        # Occurrence counter: 1-based per contiguous FXY1 run (increments when
        # the same FXY1 reappears after an interruption by another FXY1).
        # Per-(FXY1, FXY2) occurrence counter: when the same FXY2 element
        # appears multiple times within a single FXY1 sequence, each gets an
        # incrementing counter (1, 2, 3, ...) matching the English reference.
        df = df.reset_index(drop=True)
        df["seq"] = df.groupby(["FXY1", "FXY2"]).cumcount() + 1
        df["Id"] = df.apply(
            lambda r: f"bufr4/{r['FXY1'][0]}/{r['FXY1'][1:3]}/{r['FXY1'][3:]}/{r['seq']}/{r['FXY2']}",
            axis=1,
        )
        df = df.drop(columns=["seq"])

        cols = ["Id", "Category", cat_col, "FXY1", title_col, subtitle_col,
                "FXY2", name_col, desc_col, note_col, "noteIDs", "Status"]
        results[cat["category"]] = df[[c for c in cols if c in df.columns]]

    doc.close()
    return results


def save_table_d(results: dict[str, pd.DataFrame], output_dir: str, lang: str) -> int:
    """Write Table D category CSVs to *output_dir*. Returns total rows saved."""
    import os
    os.makedirs(output_dir, exist_ok=True)
    total = 0
    for cat_no, df in sorted(results.items()):
        path = os.path.join(output_dir, f"BUFR_TableD_{lang}_{cat_no}.csv")
        df.to_csv(path, index=False)
        total += len(df)
        print(f"  Category {cat_no}: {len(df)} rows → {path}")
    return total
