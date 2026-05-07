"""Notes processors — ABC for populating noteIDs and Note columns.

Notes processors take ID-validated CSVs and enrich them with noteIDs
and translated Note_{lang} columns by matching against an English reference.

To implement a new notes processor:

    from wmo.notes import BaseNotesProcessor
    from wmo.registry import registry

    class CustomNotesProcessor(BaseNotesProcessor):
        def populate(self, input_dir, ref_dir, output_dir, lang):
            ...  # return summary dict

    registry.register_notes("grib", CustomNotesProcessor)
"""

from __future__ import annotations

from abc import ABC, abstractmethod


class BaseNotesProcessor(ABC):
    """Abstract base class for notes population."""

    @abstractmethod
    def populate(
        self,
        input_dir: str,
        ref_dir: str,
        output_dir: str,
        lang: str,
    ) -> dict:
        """Populate noteIDs and Note_{lang} in all CSVs.

        Parameters
        ----------
        input_dir : str
            Directory with ID-validated language CSVs.
        ref_dir : str
            English reference CSV directory.
        output_dir : str
            Directory to write enriched CSVs.
        lang : str
            Two-letter language code.

        Returns
        -------
        dict
            Per-file summary: {filename: {matched, populated, total, ref}}
        """
        ...
