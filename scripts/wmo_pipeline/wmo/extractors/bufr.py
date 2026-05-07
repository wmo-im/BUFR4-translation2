"""BUFR table extractors — adapters to the built-in BUFR engine.

Each class wraps the corresponding extraction function from the
built-in engine (``wmo._bufr_engine``) and exposes it through the
``BaseExtractor`` interface.  The underlying extraction logic
(position-based PDF parsing via pymupdf) is unchanged.

Auto-registers all five BUFR extractors with the global registry on import.
"""

from __future__ import annotations

import pathlib
from typing import Union

import pandas as pd

from wmo import ensure_bufr_engine
from wmo.extractors import BaseExtractor
from wmo.registry import registry

# Ensure engine modules are importable
ensure_bufr_engine()

from extract.extract_abc import (  # noqa: E402
    extract_table_a as _extract_table_a,
    extract_table_b as _extract_table_b,
    extract_table_c as _extract_table_c,
    extract_codeflag as _extract_codeflag,
    save_table_b as _save_table_b,
    save_codeflag as _save_codeflag,
)
from extract.extract_d import (  # noqa: E402
    extract_table_d as _extract_table_d,
    save_table_d as _save_table_d,
)


class BufrTableAExtractor(BaseExtractor):
    """BUFR Table A — data categories (markdown-based extraction)."""

    standard = "bufr"
    table_type = "table_a"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        return _extract_table_a(pdf_path, start_page, end_page, lang, lang_config)

    def save(self, results, output_dir, lang):
        if isinstance(results, pd.DataFrame) and not results.empty:
            out = pathlib.Path(output_dir) / f"BUFR_TableA_{lang}.csv"
            results.to_csv(out, index=False)
            return len(results)
        return 0


class BufrTableBExtractor(BaseExtractor):
    """BUFR Table B — element descriptors (position-based, multi-file by class)."""

    standard = "bufr"
    table_type = "table_b"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        return _extract_table_b(pdf_path, start_page, end_page, lang, lang_config)

    def save(self, results, output_dir, lang):
        return _save_table_b(results, str(output_dir), lang)


class BufrTableCExtractor(BaseExtractor):
    """BUFR Table C — data description operators (position-based, single-file)."""

    standard = "bufr"
    table_type = "table_c"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        return _extract_table_c(pdf_path, start_page, end_page, lang, lang_config)

    def save(self, results, output_dir, lang):
        if isinstance(results, pd.DataFrame) and not results.empty:
            out = pathlib.Path(output_dir) / f"BUFRCREX_TableC_{lang}.csv"
            results.to_csv(out, index=False)
            return len(results)
        return 0


class BufrTableDExtractor(BaseExtractor):
    """BUFR Table D — common sequences (position-based, multi-file by category)."""

    standard = "bufr"
    table_type = "table_d"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        return _extract_table_d(pdf_path, start_page, end_page, lang, lang_config)

    def save(self, results, output_dir, lang):
        return _save_table_d(results, str(output_dir), lang)


class BufrCodeFlagExtractor(BaseExtractor):
    """BUFR CodeFlag — code/flag tables (position-based, multi-file by class)."""

    standard = "bufr"
    table_type = "codeflag"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        return _extract_codeflag(pdf_path, start_page, end_page, lang, lang_config)

    def save(self, results, output_dir, lang):
        return _save_codeflag(results, str(output_dir), lang)


# ── Auto-register all BUFR extractors ─────────────────────────────────────────

_BUFR_EXTRACTORS = [
    BufrTableAExtractor,
    BufrTableBExtractor,
    BufrTableCExtractor,
    BufrTableDExtractor,
    BufrCodeFlagExtractor,
]

for _cls in _BUFR_EXTRACTORS:
    registry.register_extractor(_cls.standard, _cls.table_type, _cls)
