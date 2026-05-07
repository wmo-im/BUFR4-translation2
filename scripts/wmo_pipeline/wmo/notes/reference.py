"""Reference notes processor — populates noteIDs and Note_{lang} columns.

Matches language CSV rows by Id to English reference CSVs, copying
noteIDs and translating Note_en to Note_{lang} using language-specific
rules.

Auto-registers as the default notes processor for all known standards.
"""

from __future__ import annotations

from wmo import ensure_bufr_engine
from wmo.notes import BaseNotesProcessor
from wmo.registry import registry

ensure_bufr_engine()

from populate_notes import populate_notes as _populate_notes  # noqa: E402


class ReferenceNotesProcessor(BaseNotesProcessor):
    """Populates notes from English reference CSVs."""

    def populate(self, input_dir, ref_dir, output_dir, lang):
        return _populate_notes(input_dir, ref_dir, output_dir, lang)


for _std in ("bufr", "grib", "cct"):
    registry.register_notes(_std, ReferenceNotesProcessor)
