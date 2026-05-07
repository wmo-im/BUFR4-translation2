"""Extractors — ABC and helpers for PDF → CSV table extraction.

To implement a new standard's extractor:

    from wmo.extractors import BaseExtractor
    from wmo.registry import registry

    class GribParameterExtractor(BaseExtractor):
        standard = "grib"
        table_type = "parameter"

        def extract(self, pdf_path, start_page, end_page, lang, lang_config):
            ...  # your extraction logic
            return {"00": df_class_00, "01": df_class_01}  # or just a DataFrame

        def save(self, results, output_dir, lang):
            ...  # save CSVs, return total row count

    # Auto-register
    registry.register_extractor("grib", "parameter", GribParameterExtractor)
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Union

import pandas as pd


class BaseExtractor(ABC):
    """Abstract base class for all table extractors.

    Subclasses must declare ``standard`` and ``table_type`` class attributes
    and implement ``extract()`` and ``save()``.
    """

    standard: str = ""      # e.g. "bufr", "grib", "cct"
    table_type: str = ""    # e.g. "table_a", "table_b", "parameter"

    @abstractmethod
    def extract(
        self,
        pdf_path: str,
        start_page: int,
        end_page: int,
        lang: str,
        lang_config: object,
    ) -> Union[pd.DataFrame, dict[str, pd.DataFrame]]:
        """Extract table data from a PDF page range.

        Parameters
        ----------
        pdf_path : str
            Path to the source PDF file.
        start_page : int
            First page to extract (0-indexed).
        end_page : int
            Last page to extract (0-indexed, inclusive).
        lang : str
            Two-letter language code.
        lang_config : object
            Language-specific configuration (LangConfig or similar).

        Returns
        -------
        pd.DataFrame or dict[str, pd.DataFrame]
            Single DataFrame for single-file tables (A, C).
            Dict of {key: DataFrame} for multi-file tables (B, D, CodeFlag),
            where key is typically the class/category number.
        """
        ...

    @abstractmethod
    def save(
        self,
        results: Union[pd.DataFrame, dict[str, pd.DataFrame]],
        output_dir: str,
        lang: str,
    ) -> int:
        """Save extraction results to CSV files.

        Parameters
        ----------
        results : DataFrame or dict of DataFrames
            Output from ``extract()``.
        output_dir : str
            Directory to write CSV files into.
        lang : str
            Language code (used in filenames).

        Returns
        -------
        int
            Total number of rows written.
        """
        ...

    @property
    def display_name(self) -> str:
        """Human-readable name for progress messages."""
        return f"{self.standard.upper()} {self.table_type}"
