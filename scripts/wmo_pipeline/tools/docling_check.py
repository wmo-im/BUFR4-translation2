"""Cross-check PDF table extraction against Docling's AI-based table detection.

Docling uses deep learning (layout analysis + table structure recognition)
to find and parse tables in PDFs.  This tool extracts a page range, runs
Docling on it, dumps the results, and optionally compares against an
existing set of CSVs (from any extraction pipeline).

Usage
-----
    # Extract and dump Docling output for pages 10–15:
    docling-check --pdf report.pdf --pages 10-15 --output-dir /tmp/docling_out/

    # Process entire PDF (no page slicing):
    docling-check --pdf report.pdf --output-dir /tmp/docling_out/

    # Compare Docling output against existing CSVs:
    docling-check --pdf report.pdf --pages 10-15 \\
        --compare-dir my_extracted_csvs/ \\
        --output-dir /tmp/docling_crosscheck/

    # WMO pipeline integration (requires wmo_pipeline installed):
    docling-check --config configs/bufr_fr_2019.yaml --table table_b \\
        --output-dir /tmp/docling_test/

Requirements
------------
    pip install docling-check          # this package
    # or: pip install pymupdf pandas docling
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

import pymupdf  # fitz
import pandas as pd


# ---------------------------------------------------------------------------
# Page extraction
# ---------------------------------------------------------------------------

def extract_pages(pdf_path: str | Path, start: int, end: int) -> Path:
    """Extract pages [start, end] (0-indexed, inclusive) into a temp PDF.

    Returns the path to the temporary PDF.  Caller is responsible for
    cleanup (or let the OS handle it via tempfile).
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    src = pymupdf.open(str(pdf_path))
    dst = pymupdf.open()
    dst.insert_pdf(src, from_page=start, to_page=end)

    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    dst.save(tmp.name)
    dst.close()
    src.close()

    print(f"  Extracted pages {start}–{end} → {tmp.name} "
          f"({end - start + 1} pages, {Path(tmp.name).stat().st_size / 1024:.0f} KB)")
    return Path(tmp.name)


# ---------------------------------------------------------------------------
# Docling conversion
# ---------------------------------------------------------------------------

def run_docling(pdf_path: str | Path):
    """Run Docling DocumentConverter on a PDF.

    Returns the Docling ConversionResult.
    First run downloads AI models (~1–2 GB).
    """
    try:
        from docling.document_converter import DocumentConverter
    except ImportError:
        print("ERROR: docling is not installed.")
        print("Install with:  pip install docling")
        sys.exit(1)

    print(f"  Running Docling on {pdf_path} ...")
    converter = DocumentConverter()
    result = converter.convert(str(pdf_path))
    print(f"  Docling conversion complete.")
    return result


def docling_tables_to_dataframes(result) -> list[pd.DataFrame]:
    """Extract tables from a Docling ConversionResult as DataFrames."""
    doc = result.document
    tables = []
    for i, table in enumerate(doc.tables):
        df = table.export_to_dataframe(doc=doc)
        tables.append(df)
        print(f"  Table {i}: {len(df)} rows × {len(df.columns)} cols")
    if not tables:
        print("  WARNING: Docling detected no tables in this PDF range.")
    return tables


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def dump_output(result, tables: list[pd.DataFrame], output_dir: Path) -> None:
    """Write Docling results: full markdown + per-table CSVs."""
    output_dir.mkdir(parents=True, exist_ok=True)

    # Full document as markdown
    md_path = output_dir / "docling_raw.md"
    md_text = result.document.export_to_markdown()
    md_path.write_text(md_text, encoding="utf-8")
    print(f"  Wrote {md_path} ({len(md_text)} chars)")

    # Per-table CSVs
    for i, df in enumerate(tables):
        csv_path = output_dir / f"docling_table_{i:02d}.csv"
        df.to_csv(csv_path, index=False, encoding="utf-8")
        print(f"  Wrote {csv_path} ({len(df)} rows)")


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def _find_compare_csvs(compare_dir: Path, glob_pattern: str | None = None) -> list[Path]:
    """Find CSVs in the compare directory.

    If glob_pattern is given, use it (e.g. "TableB_fr*.csv").
    Otherwise, return all .csv files.
    """
    compare_dir = Path(compare_dir)
    if not compare_dir.exists():
        return []

    pattern = glob_pattern or "*.csv"
    return sorted(compare_dir.glob(pattern))


def _load_csvs(csv_paths: list[Path]) -> pd.DataFrame:
    """Load and concatenate CSVs into a single DataFrame."""
    if not csv_paths:
        return pd.DataFrame()
    dfs = []
    for p in csv_paths:
        df = pd.read_csv(p, dtype=str, keep_default_na=False)
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)


def compare_with_csvs(
    docling_tables: list[pd.DataFrame],
    compare_dir: str | Path,
    output_dir: Path,
    glob_pattern: str | None = None,
    label: str = "existing pipeline",
) -> None:
    """Compare Docling output vs existing CSVs and write a report.

    Compares:
    1. Row counts (total and per-table)
    2. Column overlap
    3. Cell-level content differences (sampled)
    """
    compare_dir = Path(compare_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_paths = _find_compare_csvs(compare_dir, glob_pattern)
    if not csv_paths:
        report = (f"# Comparison Report\n\n"
                  f"No CSVs found in `{compare_dir}`"
                  f"{f' matching `{glob_pattern}`' if glob_pattern else ''}.\n")
        (output_dir / "comparison_report.md").write_text(report, encoding="utf-8")
        print(f"  WARNING: No CSVs found for comparison.")
        return

    pipeline_df = _load_csvs(csv_paths)

    if docling_tables:
        docling_df = pd.concat(docling_tables, ignore_index=True)
    else:
        docling_df = pd.DataFrame()

    lines: list[str] = []
    lines.append("# Docling vs Existing Extraction — Comparison Report\n")
    lines.append(f"- **Comparison source:** `{compare_dir}` ({label})")
    lines.append(f"- **CSVs loaded:** {len(csv_paths)} files")
    lines.append(f"- **Docling tables detected:** {len(docling_tables)}\n")

    # --- Row counts ---
    lines.append("## Row Counts\n")
    lines.append(f"| Source | Total Rows |")
    lines.append(f"|--------|-----------|")
    lines.append(f"| {label} | {len(pipeline_df)} |")
    lines.append(f"| Docling | {len(docling_df)} |")
    diff = len(docling_df) - len(pipeline_df)
    if diff > 0:
        lines.append(f"\nDocling found **{diff} more rows** than {label}.")
    elif diff < 0:
        lines.append(f"\n{label} found **{-diff} more rows** than Docling.")
    else:
        lines.append(f"\nRow counts match exactly.")

    # --- Per-table breakdown (Docling) ---
    if len(docling_tables) > 1:
        lines.append("\n### Docling per-table breakdown\n")
        lines.append("| Table # | Rows | Columns |")
        lines.append("|---------|------|---------|")
        for i, df in enumerate(docling_tables):
            col_names = [str(c) for c in df.columns[:5]]
            lines.append(f"| {i} | {len(df)} | {', '.join(col_names)}{'...' if len(df.columns) > 5 else ''} |")

    # --- Source files breakdown ---
    lines.append(f"\n### {label} per-file breakdown\n")
    lines.append("| File | Rows |")
    lines.append("|------|------|")
    for p in csv_paths:
        n = len(pd.read_csv(p, dtype=str, keep_default_na=False))
        lines.append(f"| `{p.name}` | {n} |")

    # --- Column comparison ---
    lines.append("\n## Column Comparison\n")
    lines.append(f"- **{label} columns:** {list(pipeline_df.columns)}")
    if not docling_df.empty:
        docling_col_names = [str(c) for c in docling_df.columns]
        lines.append(f"- **Docling columns:** {docling_col_names}")
        common = set(str(c) for c in pipeline_df.columns) & set(docling_col_names)
        if common:
            lines.append(f"- **Common columns:** {sorted(common)}")
        else:
            lines.append("- **No column names in common** (Docling uses auto-detected headers)")
    else:
        lines.append("- Docling produced no table data.")

    # --- Content sample (first 10 rows side by side) ---
    lines.append("\n## Content Sample (first 10 rows)\n")
    lines.append(f"### {label}\n")
    lines.append("```")
    lines.append(pipeline_df.head(10).to_string(index=False))
    lines.append("```\n")
    if not docling_df.empty:
        lines.append("### Docling\n")
        lines.append("```")
        lines.append(docling_df.head(10).to_string(index=False))
        lines.append("```\n")

    # --- Text content overlap ---
    lines.append("## Text Content Overlap\n")
    if not docling_df.empty:
        pipeline_texts = set()
        for col in pipeline_df.columns:
            pipeline_texts.update(
                pipeline_df[col].astype(str).str.strip().values
            )
        pipeline_texts.discard("")

        docling_texts = set()
        for col in docling_df.columns:
            docling_texts.update(
                docling_df[col].astype(str).str.strip().values
            )
        docling_texts.discard("")

        if pipeline_texts:
            found = pipeline_texts & docling_texts
            pct = 100 * len(found) / len(pipeline_texts)
            lines.append(
                f"- {len(found)}/{len(pipeline_texts)} unique cell values from {label} "
                f"also appear in Docling output ({pct:.1f}%)"
            )

            missing = pipeline_texts - docling_texts
            if missing:
                sample = sorted(missing)[:20]
                lines.append(f"\n### Sample values in {label} but NOT in Docling ({len(missing)} total):\n")
                for v in sample:
                    lines.append(f"- `{v}`")

            extra = docling_texts - pipeline_texts
            if extra:
                sample = sorted(extra)[:20]
                lines.append(f"\n### Sample values in Docling but NOT in {label} ({len(extra)} total):\n")
                for v in sample:
                    lines.append(f"- `{v}`")
        else:
            lines.append(f"- No text cells in {label} output to compare.")
    else:
        lines.append("- Docling produced no tables; cannot compare text content.")

    report_text = "\n".join(lines) + "\n"
    report_path = output_dir / "comparison_report.md"
    report_path.write_text(report_text, encoding="utf-8")
    print(f"  Wrote comparison report: {report_path}")


# ---------------------------------------------------------------------------
# Config loading helpers (optional wmo_pipeline integration)
# ---------------------------------------------------------------------------

def _load_wmo_config(config_path: str, table_type: str):
    """Load a wmo_pipeline YAML config. Returns (pdf_path, start, end, lang).

    Raises ImportError if wmo_pipeline is not available.
    """
    # Try importing from wmo_pipeline
    try:
        from wmo.config import load_config
    except ImportError:
        # Try adding parent directory to path (for running from tools/)
        tools_dir = Path(__file__).resolve().parent
        pipeline_dir = tools_dir.parent
        sys.path.insert(0, str(pipeline_dir))
        try:
            from wmo.config import load_config
        except ImportError:
            raise ImportError(
                "wmo_pipeline is not installed or not on PYTHONPATH.\n"
                "The --config option requires the wmo_pipeline package.\n"
                "Use --pdf + --pages instead for standalone operation."
            )

    cfg = load_config(config_path)

    # Normalize table key
    table_key = table_type.lower().strip()
    aliases = {
        "a": "table_a", "b": "table_b", "c": "table_c",
        "d": "table_d", "codeflag": "codeflag",
    }
    table_key = aliases.get(table_key, table_key)

    if table_key not in cfg.page_ranges:
        available = list(cfg.page_ranges.keys())
        raise ValueError(
            f"Table '{table_type}' not found in config page_ranges. "
            f"Available: {available}"
        )

    start, end = cfg.page_ranges[table_key]
    return cfg.pdf_path, start, end, cfg.lang


def _parse_pages(pages_str: str) -> tuple[int, int]:
    """Parse 'START-END' or 'START' into (start, end) integers."""
    if "-" in pages_str:
        parts = pages_str.split("-", 1)
        return int(parts[0]), int(parts[1])
    else:
        p = int(pages_str)
        return p, p


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        prog="docling-check",
        description="Run Docling AI table detection on a PDF and optionally compare against existing CSVs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:

  # Extract tables from pages 10–20 of a PDF:
  docling-check --pdf report.pdf --pages 10-20 --output-dir /tmp/out/

  # Process entire PDF:
  docling-check --pdf report.pdf --output-dir /tmp/out/

  # Compare Docling output against your own CSVs:
  docling-check --pdf report.pdf --pages 10-20 \\
      --compare-dir my_csvs/ --output-dir /tmp/crosscheck/

  # Filter which CSVs to compare (glob pattern):
  docling-check --pdf report.pdf --pages 10-20 \\
      --compare-dir my_csvs/ --compare-glob "TableB*.csv" \\
      --output-dir /tmp/crosscheck/

  # WMO pipeline integration (requires wmo_pipeline):
  docling-check --config configs/bufr_fr_2019.yaml --table table_b \\
      --output-dir /tmp/out/
        """,
    )

    # Source
    source = parser.add_argument_group("Source (choose one)")
    source.add_argument("--pdf", help="Path to source PDF")
    source.add_argument("--pages", help="Page range, 0-indexed inclusive: START-END (e.g. 10-20). "
                        "Omit to process entire PDF.")
    source.add_argument("--config", help="Path to wmo_pipeline YAML config (requires wmo_pipeline)")
    source.add_argument("--table", help="Table type for --config mode (e.g. table_b, codeflag)")

    # Output
    parser.add_argument("--output-dir", required=True, help="Directory for Docling output")

    # Comparison
    compare = parser.add_argument_group("Comparison (optional)")
    compare.add_argument("--compare-dir", help="Directory with existing CSVs to compare against")
    compare.add_argument("--compare-glob", help="Glob pattern to filter CSVs (default: *.csv)")
    compare.add_argument("--compare-label", default="existing extraction",
                         help="Label for the comparison source in the report")

    # Options
    parser.add_argument("--max-pages", type=int, default=None,
                        help="Limit number of pages to process (for quick tests)")

    args = parser.parse_args()

    # --- Resolve source ---
    use_page_slice = True

    if args.config:
        if not args.table:
            parser.error("--table is required when using --config")
        try:
            pdf_path, start, end, lang = _load_wmo_config(args.config, args.table)
        except ImportError as e:
            parser.error(str(e))
    elif args.pdf:
        pdf_path = str(Path(args.pdf).resolve())
        if args.pages:
            start, end = _parse_pages(args.pages)
        else:
            use_page_slice = False
            start = end = 0  # unused
        lang = "unknown"
    else:
        parser.error("Provide either --pdf (with optional --pages), or --config + --table")

    if not Path(pdf_path).exists():
        print(f"ERROR: PDF not found: {pdf_path}")
        sys.exit(1)

    # Apply max-pages limit
    if use_page_slice and args.max_pages and (end - start + 1) > args.max_pages:
        end = start + args.max_pages - 1
        print(f"  Limiting to {args.max_pages} pages: {start}–{end}")

    output_dir = Path(args.output_dir)

    # --- Run ---
    print(f"\n{'='*60}")
    print(f"Docling Cross-Check")
    print(f"{'='*60}")
    print(f"  PDF:    {pdf_path}")
    if use_page_slice:
        print(f"  Pages:  {start}–{end} ({end - start + 1} pages)")
    else:
        print(f"  Pages:  all")
    print(f"  Output: {output_dir}")
    if args.compare_dir:
        print(f"  Compare: {args.compare_dir}")
        if args.compare_glob:
            print(f"  Glob:    {args.compare_glob}")
    print()

    # Step 1: Extract pages (or use full PDF)
    if use_page_slice:
        print("[1/3] Extracting pages ...")
        working_pdf = extract_pages(pdf_path, start, end)
        cleanup_pdf = True
    else:
        print("[1/3] Using full PDF (no page extraction) ...")
        working_pdf = Path(pdf_path)
        cleanup_pdf = False

    # Step 2: Run Docling
    print("[2/3] Running Docling ...")
    result = run_docling(working_pdf)
    tables = docling_tables_to_dataframes(result)

    # Step 3: Dump output
    print("[3/3] Writing output ...")
    dump_output(result, tables, output_dir)

    # Optional: Compare
    if args.compare_dir:
        print("\n[bonus] Comparing with existing CSVs ...")
        compare_with_csvs(
            tables, args.compare_dir, output_dir,
            glob_pattern=args.compare_glob,
            label=args.compare_label,
        )

    # Cleanup temp PDF
    if cleanup_pdf:
        try:
            working_pdf.unlink()
        except OSError:
            pass

    print(f"\nDone. Output in {output_dir}/")


if __name__ == "__main__":
    main()
