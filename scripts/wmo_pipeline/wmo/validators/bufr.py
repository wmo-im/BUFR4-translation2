"""BUFR ID validator — adapter to the built-in BUFR engine.

Validates and fixes the Id column in BUFR CSVs using the WMO-standard
ID format:
  - Table A:    bufr4/a/{CodeFigure}
  - Table B:    bufr4/{F}/{XX}/{YYY}
  - Table C:    bufr4/{FXY}
  - Table D:    bufr4/3/{XX}/{YYY}/{occ}/{FXY2:06d}
  - CodeFlag:   bufr4/{F}/{XX}/{YYY}/{cf}

Auto-registers with the global registry on import.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from wmo import ensure_bufr_engine
from wmo.validators import BaseValidator
from wmo.registry import registry

ensure_bufr_engine()

from id_utils import fix_file as _fix_file, fix_directory as _fix_directory  # noqa: E402


class BufrIdValidator(BaseValidator):
    """BUFR ID validator using the existing id_utils module."""

    def validate_file(self, path, output_dir=None):
        return _fix_file(path, output_dir)

    def validate_directory(self, input_dir, output_dir):
        return _fix_directory(input_dir, output_dir)


registry.register_validator("bufr", BufrIdValidator)
