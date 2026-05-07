"""Reference aligner — compares language CSVs against English reference.

Works for any standard that uses the same (Id, columns) CSV layout.
Currently used by BUFR; can be shared with GRIB/CCT if they follow
the same pattern.

Auto-registers as the default aligner for all known standards.
"""

from __future__ import annotations

from pathlib import Path

from wmo import ensure_bufr_engine
from wmo.aligners import BaseAligner
from wmo.registry import registry

ensure_bufr_engine()

from align.align_to_reference import (  # noqa: E402
    align_directory as _align_directory,
    write_summary_report as _write_summary_report,
)


class ReferenceAligner(BaseAligner):
    """Aligns language CSVs against an English reference by Id column."""

    def align_directory(self, lang_dir, ref_dir, output_dir, lang):
        return _align_directory(
            Path(lang_dir), Path(ref_dir), Path(output_dir), lang
        )

    def write_report(self, summary, report_path):
        _write_summary_report(summary, Path(report_path))


# Register for all standards (the aligner is standard-agnostic)
for _std in ("bufr", "grib", "cct"):
    registry.register_aligner(_std, ReferenceAligner)
