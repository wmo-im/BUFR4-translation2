"""WMO Table Pipeline — modular framework for extracting WMO code tables.

Supports BUFR, GRIB, and CCT standards across EN/FR/ES/RU languages.
BUFR extraction engine is built-in (wmo._bufr_engine).

Usage
-----
    from wmo.config import load_config
    from wmo.runner import run_pipeline

    cfg = load_config("configs/bufr_fr_2019.yaml")
    run_pipeline(cfg)
"""

from __future__ import annotations

from pathlib import Path

__version__ = "1.0.0"

PACKAGE_ROOT = Path(__file__).resolve().parent          # wmo/
PIPELINE_ROOT = PACKAGE_ROOT.parent                     # wmo_pipeline/


def ensure_bufr_engine() -> None:
    """Ensure the built-in BUFR engine's modules are importable.

    Imports ``wmo._bufr_engine`` which adds its directory to sys.path,
    enabling the engine's internal bare imports to resolve.

    Called lazily by BUFR adapter modules only when BUFR extraction is
    actually requested.
    """
    import wmo._bufr_engine  # noqa: F401 — triggers sys.path setup
