"""Pipeline runner — orchestrates extract → validate → align → notes.

This is the brain of the framework.  Given a ``JobConfig``, it:

1. Discovers registered plugins via ``auto_discover()``.
2. Resolves which tables to process.
3. Runs each requested step in sequence, delegating to the right
   plugin class from the registry.

Usage
-----
    from wmo.config import load_config
    from wmo.runner import run_pipeline

    cfg = load_config("configs/bufr_fr_2019.yaml")
    run_pipeline(cfg)
"""

from __future__ import annotations

import shutil
import time
from pathlib import Path

import pandas as pd

from wmo.config import JobConfig
from wmo.registry import registry, auto_discover


def run_pipeline(config: JobConfig) -> dict:
    """Execute the full pipeline as specified in *config*.

    Parameters
    ----------
    config : JobConfig
        Loaded and validated job configuration.

    Returns
    -------
    dict
        Summary with per-step results and timing.
    """
    auto_discover()

    print("=" * 70)
    print(f"  WMO Table Pipeline — {config.name}")
    print(f"  Standard: {config.standard.upper()}  |  Language: {config.lang.upper()}"
          f"  |  Edition: {config.edition}")
    print("=" * 70)
    print(f"  Steps:  {config.steps}")
    print(f"  Tables: {config.tables}")
    print(f"  Output: {config.output_dir}")
    print("=" * 70)

    # Resolve table list
    all_tables = registry.get_all_table_types(config.standard)
    if not all_tables:
        print(f"\nERROR: No extractors registered for standard '{config.standard}'.")
        print(f"Registered standards: {registry.list_standards()}")
        return {"error": f"Unknown standard: {config.standard}"}

    tables = config.resolve_tables(all_tables)
    print(f"\nResolved tables: {tables}")

    # Set up output directory structure
    lang = config.lang
    output_root = Path(config.output_dir)
    lang_dir = output_root / lang
    raw_dir = lang_dir / "raw"
    fixed_dir = lang_dir / "fixed"
    aligned_dir = lang_dir / "aligned"
    final_dir = lang_dir / "final"
    id_report = lang_dir / "id_report.csv"
    align_report = lang_dir / "alignment_report.md"

    summary: dict = {"steps": {}, "total_time": 0.0}
    pipeline_start = time.time()

    # ── Step 1: Extract ───────────────────────────────────────────────────────
    if "extract" in config.steps:
        step_result = _run_extraction(config, tables, raw_dir)
        summary["steps"]["extract"] = step_result
        source_dir = raw_dir
    else:
        source_dir = Path(config.input_dir) if config.input_dir else raw_dir
        if not source_dir.exists():
            print(f"\nWARNING: Source directory {source_dir} does not exist.")
        print(f"\n--- Step 1: Extract SKIPPED (using {source_dir}) ---")

    # ── Step 2: Validate IDs ─────────────────────────────────────────────────
    if "validate" in config.steps:
        step_result = _run_validation(config, source_dir, fixed_dir, id_report)
        summary["steps"]["validate"] = step_result
    else:
        print("\n--- Step 2: Validate SKIPPED ---")
        # If validation is skipped, use source_dir as the "fixed" dir for later steps
        if not fixed_dir.exists() and source_dir.exists():
            fixed_dir = source_dir

    # ── Step 3: Align ─────────────────────────────────────────────────────────
    if "align" in config.steps:
        step_result = _run_alignment(config, fixed_dir, aligned_dir, align_report)
        summary["steps"]["align"] = step_result
    else:
        print("\n--- Step 3: Align SKIPPED ---")

    # ── Step 4: Notes ─────────────────────────────────────────────────────────
    if "notes" in config.steps:
        step_result = _run_notes(config, fixed_dir, final_dir)
        summary["steps"]["notes"] = step_result
    else:
        print("\n--- Step 4: Notes SKIPPED ---")

    # ── Optional: Copy to destination ─────────────────────────────────────────
    if config.copy_to:
        _copy_results(lang_dir, config.copy_to)

    # ── Summary ───────────────────────────────────────────────────────────────
    total_time = time.time() - pipeline_start
    summary["total_time"] = total_time

    print("\n" + "=" * 70)
    print(f"  Pipeline complete in {total_time:.1f}s")
    print(f"  Outputs: {lang_dir}")
    if config.copy_to:
        print(f"  Copied to: {config.copy_to}")
    print("=" * 70)

    return summary


# ── Step implementations ──────────────────────────────────────────────────────

def _run_extraction(
    config: JobConfig,
    tables: list[str],
    raw_dir: Path,
) -> dict:
    """Step 1: Extract tables from PDF."""
    print(f"\n{'='*50}")
    print(f"  Step 1: EXTRACT ({config.standard.upper()}, {config.lang.upper()})")
    print(f"{'='*50}")

    start = time.time()
    raw_dir.mkdir(parents=True, exist_ok=True)

    # Load language config
    from wmo import ensure_bufr_engine
    ensure_bufr_engine()
    from lang_config import get_lang_config
    lang_config = get_lang_config(config.lang, config.lang_overrides)

    results = {}
    for i, table_type in enumerate(tables, 1):
        page_range = config.page_ranges.get(table_type)
        if page_range is None:
            print(f"\n  [{i}/{len(tables)}] {table_type}: no page range → skipping")
            continue

        extractor_cls = registry.get_extractor(config.standard, table_type)
        if extractor_cls is None:
            print(f"\n  [{i}/{len(tables)}] {table_type}: no extractor registered → skipping")
            continue

        extractor = extractor_cls()
        start_page, end_page = page_range

        print(f"\n  [{i}/{len(tables)}] Extracting {extractor.display_name} "
              f"(pages {start_page}–{end_page})...")

        data = extractor.extract(config.pdf_path, start_page, end_page,
                                 config.lang, lang_config)

        if data is None or (isinstance(data, pd.DataFrame) and data.empty):
            print(f"    → No rows extracted")
            results[table_type] = 0
            continue

        total = extractor.save(data, str(raw_dir), config.lang)
        print(f"    → {total} rows saved")
        results[table_type] = total

    elapsed = time.time() - start
    print(f"\n  Extraction complete in {elapsed:.1f}s")
    return {"tables": results, "time": elapsed}


def _run_validation(
    config: JobConfig,
    source_dir: Path,
    fixed_dir: Path,
    id_report: Path,
) -> dict:
    """Step 2: Validate and fix IDs."""
    print(f"\n{'='*50}")
    print(f"  Step 2: VALIDATE IDs")
    print(f"{'='*50}")

    start = time.time()
    validator_cls = registry.get_validator(config.standard)
    if validator_cls is None:
        print(f"  No validator registered for '{config.standard}'. Skipping.")
        return {"skipped": True}

    validator = validator_cls()
    fixed_dir.mkdir(parents=True, exist_ok=True)
    all_changes, file_count = validator.validate_directory(source_dir, fixed_dir)

    # Write report
    report_df = pd.DataFrame(
        all_changes,
        columns=["file", "row_index", "old_id", "new_id", "method"],
    )
    id_report.parent.mkdir(parents=True, exist_ok=True)
    report_df.to_csv(id_report, index=False)

    n_ok = sum(1 for c in all_changes if c["method"] != "unresolved")
    n_bad = sum(1 for c in all_changes if c["method"] == "unresolved")

    elapsed = time.time() - start
    print(f"\n  Validation complete: {file_count} files, {n_ok} fixed, "
          f"{n_bad} unresolved ({elapsed:.1f}s)")
    print(f"  Report: {id_report}")

    return {"files": file_count, "fixed": n_ok, "unresolved": n_bad, "time": elapsed}


def _run_alignment(
    config: JobConfig,
    fixed_dir: Path,
    aligned_dir: Path,
    report_path: Path,
) -> dict:
    """Step 3: Align against English reference."""
    print(f"\n{'='*50}")
    print(f"  Step 3: ALIGN vs reference")
    print(f"{'='*50}")

    if not config.ref_dir:
        print("  No ref_dir configured. Skipping alignment.")
        return {"skipped": True}

    start = time.time()
    aligner_cls = registry.get_aligner(config.standard)
    if aligner_cls is None:
        print(f"  No aligner registered for '{config.standard}'. Skipping.")
        return {"skipped": True}

    ref_dir = Path(config.ref_dir)
    if not ref_dir.exists():
        print(f"  Reference directory not found: {ref_dir}. Skipping.")
        return {"skipped": True, "reason": "ref_dir not found"}

    aligner = aligner_cls()
    aligned_dir.mkdir(parents=True, exist_ok=True)
    summary = aligner.align_directory(fixed_dir, ref_dir, aligned_dir, config.lang)
    aligner.write_report(summary, report_path)

    elapsed = time.time() - start
    print(f"\n  Alignment complete ({elapsed:.1f}s)")
    return {"summary": summary, "time": elapsed}


def _run_notes(
    config: JobConfig,
    fixed_dir: Path,
    final_dir: Path,
) -> dict:
    """Step 4: Populate notes from English reference."""
    print(f"\n{'='*50}")
    print(f"  Step 4: NOTES population")
    print(f"{'='*50}")

    if not config.ref_dir:
        print("  No ref_dir configured. Skipping notes.")
        return {"skipped": True}

    if config.lang == "en":
        print("  Language is 'en' — notes population not applicable. Skipping.")
        return {"skipped": True, "reason": "english"}

    start = time.time()
    notes_cls = registry.get_notes(config.standard)
    if notes_cls is None:
        print(f"  No notes processor registered for '{config.standard}'. Skipping.")
        return {"skipped": True}

    processor = notes_cls()
    summary = processor.populate(
        str(fixed_dir), config.ref_dir, str(final_dir), config.lang
    )

    total_populated = sum(s.get("populated", 0) for s in summary.values())
    total_rows = sum(s.get("total", 0) for s in summary.values())

    elapsed = time.time() - start
    print(f"\n  Notes complete: {total_populated} notes across {total_rows} rows ({elapsed:.1f}s)")

    return {"populated": total_populated, "total_rows": total_rows, "time": elapsed}


def _copy_results(source_dir: Path, dest_dir: str) -> None:
    """Copy final results to a destination directory."""
    dest = Path(dest_dir)
    dest.mkdir(parents=True, exist_ok=True)

    # Copy the most complete version: final/ > fixed/ > raw/
    for subdir in ("final", "fixed", "raw"):
        src = source_dir / subdir
        if src.exists():
            target = dest
            print(f"\n  Copying {src} → {target}")
            for csv_file in sorted(src.glob("*.csv")):
                shutil.copy2(csv_file, target / csv_file.name)
            break
