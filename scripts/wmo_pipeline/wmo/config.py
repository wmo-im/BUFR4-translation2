"""Job configuration — loads YAML/JSON into a typed JobConfig object.

A single YAML file describes everything needed for one pipeline run:
source PDF, page ranges, language, tables to process, steps to run,
reference directory for alignment, and output location.

Usage
-----
    from wmo.config import load_config
    cfg = load_config("configs/bufr_fr_2019.yaml")
    print(cfg.standard, cfg.lang, cfg.tables)
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class JobConfig:
    """Immutable snapshot of a pipeline run configuration.

    Attributes
    ----------
    name : str
        Human-readable job name (for logging/reports).
    standard : str
        WMO standard: "bufr", "grib", or "cct".
    lang : str
        Two-letter language code: "en", "fr", "es", "ru".
    edition : str
        Source document edition (e.g. "2019", "2019+2021").
    pdf_path : str
        Absolute path to the source PDF.
    page_ranges : dict[str, list[int]]
        Mapping from table type key to [start_page, end_page] (0-indexed).
        Example: {"table_a": [252, 253], "table_b": [256, 405], ...}
    ref_dir : str | None
        Path to English reference CSV directory (for alignment/notes).
    output_dir : str
        Root output directory. Subdirs created automatically.
    steps : list[str]
        Pipeline steps to execute, in order.
        Valid steps: "extract", "validate", "align", "notes".
    tables : list[str]
        Table types to process. Use ["all"] for all tables in the standard.
        BUFR examples: ["table_a", "table_b", "table_c", "table_d", "codeflag"]
    lang_overrides : dict | None
        Optional overrides for LangConfig fields (column bounds, patterns, etc.).
    input_dir : str | None
        Pre-extracted CSV directory (alternative to PDF extraction).
    copy_to : str | None
        Optional directory to copy final results to after pipeline completes.
    """

    name: str
    standard: str
    lang: str
    edition: str = ""
    pdf_path: str = ""
    page_ranges: dict[str, list[int]] = field(default_factory=dict)
    ref_dir: str | None = None
    output_dir: str = ""
    steps: list[str] = field(default_factory=lambda: ["extract", "validate", "align", "notes"])
    tables: list[str] = field(default_factory=lambda: ["all"])
    lang_overrides: dict | None = None
    input_dir: str | None = None
    copy_to: str | None = None

    def __post_init__(self) -> None:
        self.standard = self.standard.lower().strip()
        self.lang = self.lang.lower().strip()
        self.steps = [s.lower().strip() for s in self.steps]
        self.tables = [t.lower().strip() for t in self.tables]

    def resolve_tables(self, all_tables: list[str]) -> list[str]:
        """Expand ["all"] into the full list of table types for this standard."""
        if "all" in self.tables:
            return all_tables
        return self.tables

    def validate(self) -> list[str]:
        """Return list of validation errors (empty = OK)."""
        errors: list[str] = []
        if not self.standard:
            errors.append("'standard' is required (bufr, grib, or cct)")
        if not self.lang:
            errors.append("'lang' is required (en, fr, es, ru)")
        if not self.output_dir:
            errors.append("'output_dir' is required")
        if "extract" in self.steps:
            if not self.pdf_path and not self.input_dir:
                errors.append("'pdf_path' or 'input_dir' required for extraction step")
            if self.pdf_path and not self.page_ranges:
                errors.append("'page_ranges' required when using pdf_path")
        if "align" in self.steps or "notes" in self.steps:
            if not self.ref_dir:
                errors.append("'ref_dir' required for align/notes steps")
        return errors


def _resolve_path(value: str, config_dir: Path) -> str:
    """Resolve a path relative to the config file's directory."""
    if not value:
        return value
    p = Path(value)
    if p.is_absolute():
        return str(p)
    return str((config_dir / p).resolve())


def load_config(path: str | Path) -> JobConfig:
    """Load a YAML or JSON config file into a JobConfig.

    Relative paths in the config are resolved relative to the config
    file's directory.

    Parameters
    ----------
    path : str | Path
        Path to a .yaml, .yml, or .json configuration file.

    Returns
    -------
    JobConfig

    Raises
    ------
    FileNotFoundError
        If the config file does not exist.
    ValueError
        If required fields are missing or invalid.
    """
    path = Path(path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    raw_text = path.read_text(encoding="utf-8")
    config_dir = path.parent

    # Parse YAML or JSON
    if path.suffix in (".yaml", ".yml"):
        try:
            import yaml
        except ImportError:
            raise ImportError(
                "PyYAML is required for YAML configs. Install: pip install pyyaml"
            )
        data: dict[str, Any] = yaml.safe_load(raw_text) or {}
    else:
        data = json.loads(raw_text)

    # Flatten nested sections (support both flat and nested YAML styles)
    flat = _flatten_config(data)

    # Resolve relative paths
    for key in ("pdf_path", "ref_dir", "output_dir", "input_dir", "copy_to"):
        if key in flat and flat[key]:
            flat[key] = _resolve_path(str(flat[key]), config_dir)

    # Normalize page_ranges keys: "a" → "table_a", "codeflag" stays
    if "page_ranges" in flat:
        flat["page_ranges"] = _normalize_page_range_keys(flat["page_ranges"])

    cfg = JobConfig(**{k: v for k, v in flat.items() if k in JobConfig.__dataclass_fields__})

    errors = cfg.validate()
    if errors:
        raise ValueError(
            f"Config validation errors in {path}:\n" + "\n".join(f"  - {e}" for e in errors)
        )

    return cfg


def _flatten_config(data: dict) -> dict:
    """Flatten nested YAML sections into a flat dict.

    Supports both flat and nested styles:
        # Flat style
        pdf_path: /path/to/pdf
        ref_dir: /path/to/ref

        # Nested style
        source:
          pdf_path: /path/to/pdf
        reference:
          dir: /path/to/ref
    """
    flat: dict[str, Any] = {}

    # Direct top-level keys
    for key in ("name", "standard", "lang", "edition", "pdf_path", "page_ranges",
                "ref_dir", "output_dir", "steps", "tables", "lang_overrides",
                "input_dir", "copy_to"):
        if key in data:
            flat[key] = data[key]

    # Nested "job" section
    if "job" in data and isinstance(data["job"], dict):
        for key in ("name", "standard", "lang", "edition"):
            if key in data["job"]:
                flat.setdefault(key, data["job"][key])

    # Nested "source" section
    if "source" in data and isinstance(data["source"], dict):
        src = data["source"]
        flat.setdefault("pdf_path", src.get("pdf") or src.get("pdf_path", ""))
        flat.setdefault("page_ranges", src.get("page_ranges", {}))
        flat.setdefault("input_dir", src.get("input_dir"))

    # Nested "reference" section
    if "reference" in data and isinstance(data["reference"], dict):
        flat.setdefault("ref_dir", data["reference"].get("dir") or data["reference"].get("ref_dir"))

    # Nested "output" section
    if "output" in data and isinstance(data["output"], dict):
        flat.setdefault("output_dir", data["output"].get("dir") or data["output"].get("output_dir", ""))
        flat.setdefault("copy_to", data["output"].get("copy_to"))

    # Nested "pipeline" section
    if "pipeline" in data and isinstance(data["pipeline"], dict):
        flat.setdefault("steps", data["pipeline"].get("steps", []))
        flat.setdefault("tables", data["pipeline"].get("tables", ["all"]))

    return flat


# Short keys → canonical table type names
_TABLE_KEY_ALIASES = {
    "a": "table_a",
    "b": "table_b",
    "c": "table_c",
    "d": "table_d",
    "table_a": "table_a",
    "table_b": "table_b",
    "table_c": "table_c",
    "table_d": "table_d",
    "codeflag": "codeflag",
    "code_flag": "codeflag",
}


def _normalize_page_range_keys(page_ranges: dict) -> dict:
    """Normalize short keys like 'a' to canonical 'table_a' form."""
    result = {}
    for key, value in page_ranges.items():
        canonical = _TABLE_KEY_ALIASES.get(key.lower().strip(), key.lower().strip())
        result[canonical] = value
    return result
