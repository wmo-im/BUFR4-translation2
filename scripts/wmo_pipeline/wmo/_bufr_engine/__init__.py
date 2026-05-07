"""BUFR extraction engine — incorporated from bufr_pipeline/.

This package contains the proven BUFR PDF extraction code.  Its modules
use bare imports (``from lang_config import ...``, ``from extract.extract_abc
import ...``).  To make those work as-is without rewriting every import,
we add this package's directory to sys.path on first import.

This is an internal package — external code should use the adapter
classes in ``wmo.extractors.bufr``, ``wmo.validators.bufr``, etc.
"""

import sys
from pathlib import Path

_ENGINE_DIR = str(Path(__file__).resolve().parent)
if _ENGINE_DIR not in sys.path:
    sys.path.insert(0, _ENGINE_DIR)
