"""Aligners — ABC for comparing extracted CSVs against a reference.

Aligners compare language-specific CSVs against an English reference,
categorizing each row as MATCH, MISSING, EXTRA, etc.  They produce
per-file alignment CSVs and a summary report.

To implement a new aligner:

    from wmo.aligners import BaseAligner
    from wmo.registry import registry

    class CustomAligner(BaseAligner):
        def align_directory(self, lang_dir, ref_dir, output_dir, lang):
            ...  # return summary list

    registry.register_aligner("grib", CustomAligner)
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path


class BaseAligner(ABC):
    """Abstract base class for reference aligners."""

    @abstractmethod
    def align_directory(
        self,
        lang_dir: Path,
        ref_dir: Path,
        output_dir: Path,
        lang: str,
    ) -> list[dict]:
        """Align all language CSVs against reference CSVs.

        Returns
        -------
        list[dict]
            Summary with per-file statistics:
            {file, ref_file, lang_rows, ref_rows, match, missing, extra, ...}
        """
        ...

    @abstractmethod
    def write_report(self, summary: list[dict], report_path: Path) -> None:
        """Write a human-readable alignment report."""
        ...
