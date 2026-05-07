"""Validators — ABC for ID validation and fixing.

Each WMO standard has its own ID format. Validators check and fix the
Id column in extracted CSVs to ensure they conform to the standard's
format specification.

To implement a new validator:

    from wmo.validators import BaseValidator
    from wmo.registry import registry

    class GribIdValidator(BaseValidator):
        def validate_directory(self, input_dir, output_dir):
            ...  # validate all CSVs, return (changes, file_count)

    registry.register_validator("grib", GribIdValidator)
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

import pandas as pd


class BaseValidator(ABC):
    """Abstract base class for ID validators."""

    @abstractmethod
    def validate_file(
        self,
        path: Path,
        output_dir: Path | None = None,
    ) -> tuple[pd.DataFrame, list[dict]]:
        """Validate and fix IDs in one CSV file.

        Parameters
        ----------
        path : Path
            Input CSV path.
        output_dir : Path | None
            If given, write the fixed CSV here.

        Returns
        -------
        (fixed_df, changes_list)
            changes_list entries: {file, row_index, old_id, new_id, method}
        """
        ...

    @abstractmethod
    def validate_directory(
        self,
        input_dir: Path,
        output_dir: Path,
    ) -> tuple[list[dict], int]:
        """Validate all CSVs in a directory.

        Returns
        -------
        (all_changes, file_count)
        """
        ...
