"""GRIB table extractors — skeleton for future implementation.

GRIB (WMO-306 Vol. I.2 Part B) defines its own set of tables:
  - Parameter tables (Code Table 4.2) — physical parameters
  - Level/Layer tables (Code Table 4.5) — vertical coordinates
  - Product definition templates (Table 4.0)
  - Data representation templates (Table 5.0)
  - Grid definition templates (Table 3.1)

To implement:
1. Subclass BaseExtractor for each GRIB table type.
2. Implement extract() and save() with GRIB-specific parsing logic.
3. Register with the global registry.
4. Add a LangConfig variant if GRIB PDFs have different column layouts.
5. Create a YAML config (e.g., configs/grib_fr_2019.yaml).

Example GRIB config YAML:
    name: "GRIB French 2019"
    standard: grib
    lang: fr
    edition: "2019"
    pdf_path: "/path/to/grib_pdf.pdf"
    page_ranges:
      parameter: [10, 50]
      level: [51, 70]
    ref_dir: "/path/to/grib_en_reference/"
    output_dir: "/path/to/output/"
    steps: [extract, validate, align]
    tables: [all]
"""

from __future__ import annotations

import pandas as pd

from wmo.extractors import BaseExtractor
from wmo.registry import registry


class GribParameterExtractor(BaseExtractor):
    """GRIB parameter table extractor — PLACEHOLDER."""

    standard = "grib"
    table_type = "parameter"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        raise NotImplementedError(
            "GRIB parameter extraction not yet implemented. "
            "See this file for implementation guide."
        )

    def save(self, results, output_dir, lang):
        raise NotImplementedError


class GribLevelExtractor(BaseExtractor):
    """GRIB level/layer table extractor — PLACEHOLDER."""

    standard = "grib"
    table_type = "level"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        raise NotImplementedError(
            "GRIB level extraction not yet implemented."
        )

    def save(self, results, output_dir, lang):
        raise NotImplementedError


class GribTemplateExtractor(BaseExtractor):
    """GRIB template table extractor — PLACEHOLDER."""

    standard = "grib"
    table_type = "template"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        raise NotImplementedError(
            "GRIB template extraction not yet implemented."
        )

    def save(self, results, output_dir, lang):
        raise NotImplementedError


# Register GRIB extractors (they'll show up in --list-plugins
# but raise NotImplementedError if actually invoked)
for _cls in (GribParameterExtractor, GribLevelExtractor, GribTemplateExtractor):
    registry.register_extractor(_cls.standard, _cls.table_type, _cls)
