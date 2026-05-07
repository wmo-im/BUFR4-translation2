"""CCT (Common Code Tables) extractors — skeleton for future implementation.

Common Code Tables (WMO-306 Vol. I.2 Part C) are shared across BUFR,
GRIB, and other WMO codes.  They include:
  - C-1:  Identification of originating/generating centre
  - C-2:  Radiosonde/sounding system codes
  - C-3:  Instrument types for wind profilers
  - C-5:  Satellite identifiers
  - C-6:  Data sub-categories
  - C-7:  Tracking techniques
  - C-11: Originating/generating sub-centres
  - C-12: Instrument types for AMDAR
  - C-13: Data sub-categories per originating centre
  - C-14: Atmospheric chemical constituents

To implement:
1. Subclass BaseExtractor for each CCT table (or one generic class).
2. CCT tables are typically simpler (two-column code→meaning).
3. Register with the global registry.
4. Create a YAML config.

Example CCT config YAML:
    name: "CCT French 2019"
    standard: cct
    lang: fr
    edition: "2019"
    pdf_path: "/path/to/cct_pdf.pdf"
    page_ranges:
      c1: [5, 20]
      c5: [30, 45]
    ref_dir: "/path/to/cct_en_reference/"
    output_dir: "/path/to/output/"
    steps: [extract, validate]
    tables: [all]
"""

from __future__ import annotations

from wmo.extractors import BaseExtractor
from wmo.registry import registry


class CctGenericExtractor(BaseExtractor):
    """Generic CCT table extractor — PLACEHOLDER.

    CCT tables share a common structure (code + meaning), so one
    extractor class may suffice for all CCT tables with table_type
    passed as a parameter.
    """

    standard = "cct"
    table_type = "generic"

    def extract(self, pdf_path, start_page, end_page, lang, lang_config):
        raise NotImplementedError(
            "CCT extraction not yet implemented. "
            "See this file for implementation guide."
        )

    def save(self, results, output_dir, lang):
        raise NotImplementedError


registry.register_extractor("cct", "generic", CctGenericExtractor)
