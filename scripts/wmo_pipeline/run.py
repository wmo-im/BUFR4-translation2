#!/usr/bin/env python3
"""WMO Table Pipeline — single-command entry point.

Usage
-----
    # Run with a YAML config (one button):
    python run.py --config configs/bufr_fr_2019.yaml

    # Override output directory:
    python run.py --config configs/bufr_fr_2019.yaml --output-dir /tmp/test/

    # List registered plugins:
    python run.py --list-plugins

    # Validate config without running:
    python run.py --config configs/bufr_fr_2019.yaml --dry-run
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Ensure wmo package is importable regardless of working directory
_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="python run.py",
        description="WMO Table Pipeline — extract, validate, align, and populate "
                    "WMO code tables from multilingual PDFs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--config", "-c",
        help="Path to a YAML/JSON configuration file",
    )
    parser.add_argument(
        "--output-dir", "-o",
        help="Override the output directory from the config file",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate config and show what would be done, without running",
    )
    parser.add_argument(
        "--list-plugins",
        action="store_true",
        help="List all registered extractor/validator/aligner/notes plugins",
    )
    parser.add_argument(
        "--steps",
        help="Override steps (comma-separated): extract,validate,align,notes",
    )

    args = parser.parse_args()

    # List plugins
    if args.list_plugins:
        from wmo.registry import registry, auto_discover
        auto_discover()
        print(registry.summary())
        return

    # Config is required for all other operations
    if not args.config:
        parser.error("--config is required (or use --list-plugins)")

    from wmo.config import load_config
    from wmo.runner import run_pipeline

    # Load config
    try:
        config = load_config(args.config)
    except (FileNotFoundError, ValueError, ImportError) as e:
        sys.exit(f"ERROR: {e}")

    # Apply overrides
    if args.output_dir:
        config.output_dir = str(Path(args.output_dir).resolve())
    if args.steps:
        config.steps = [s.strip() for s in args.steps.split(",")]

    # Dry run
    if args.dry_run:
        print("=== DRY RUN ===")
        print(f"  Name:     {config.name}")
        print(f"  Standard: {config.standard}")
        print(f"  Language: {config.lang}")
        print(f"  Edition:  {config.edition}")
        print(f"  PDF:      {config.pdf_path}")
        print(f"  Pages:    {config.page_ranges}")
        print(f"  Ref dir:  {config.ref_dir}")
        print(f"  Output:   {config.output_dir}")
        print(f"  Steps:    {config.steps}")
        print(f"  Tables:   {config.tables}")
        print(f"  Copy to:  {config.copy_to}")
        print("Config is valid. Pipeline would proceed with the above settings.")
        return

    # Run the pipeline
    run_pipeline(config)


if __name__ == "__main__":
    main()
