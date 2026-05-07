"""Plugin registry — maps (standard, table_type) to handler classes.

The registry is the extensibility backbone of the framework.  When adding
a new standard (GRIB, CCT) or a new table type, you:

1. Create an extractor class implementing ``BaseExtractor``.
2. Call ``registry.register_extractor("grib", "parameter", GribParameterExtractor)``.
3. Optionally create a validator implementing ``BaseValidator``.

The runner queries the registry at runtime to find the right handler
for each (standard, table_type) combination.

Auto-registration
-----------------
BUFR extractors register themselves when ``wmo.extractors.bufr`` is imported.
To trigger auto-registration for all known standards, call ``auto_discover()``.
"""

from __future__ import annotations

from typing import Type


class Registry:
    """Central registry for pipeline plugin classes.

    Stores four separate registries:
    - extractors:  (standard, table_type) → ExtractorClass
    - validators:  standard → ValidatorClass
    - aligners:    standard → AlignerClass
    - notes:       standard → NotesProcessorClass
    """

    def __init__(self) -> None:
        self._extractors: dict[tuple[str, str], Type] = {}
        self._validators: dict[str, Type] = {}
        self._aligners: dict[str, Type] = {}
        self._notes: dict[str, Type] = {}
        self._standard_tables: dict[str, list[str]] = {}

    # ── Extractors ────────────────────────────────────────────────────────────

    def register_extractor(
        self,
        standard: str,
        table_type: str,
        cls: Type,
    ) -> None:
        """Register an extractor class for a (standard, table_type) pair."""
        key = (standard.lower(), table_type.lower())
        self._extractors[key] = cls
        # Track known table types per standard
        tables = self._standard_tables.setdefault(standard.lower(), [])
        if table_type.lower() not in tables:
            tables.append(table_type.lower())

    def get_extractor(self, standard: str, table_type: str) -> Type | None:
        """Look up an extractor class by (standard, table_type)."""
        return self._extractors.get((standard.lower(), table_type.lower()))

    def get_all_table_types(self, standard: str) -> list[str]:
        """Return all registered table types for a standard."""
        return list(self._standard_tables.get(standard.lower(), []))

    # ── Validators ────────────────────────────────────────────────────────────

    def register_validator(self, standard: str, cls: Type) -> None:
        self._validators[standard.lower()] = cls

    def get_validator(self, standard: str) -> Type | None:
        return self._validators.get(standard.lower())

    # ── Aligners ──────────────────────────────────────────────────────────────

    def register_aligner(self, standard: str, cls: Type) -> None:
        self._aligners[standard.lower()] = cls

    def get_aligner(self, standard: str) -> Type | None:
        return self._aligners.get(standard.lower())

    # ── Notes processors ──────────────────────────────────────────────────────

    def register_notes(self, standard: str, cls: Type) -> None:
        self._notes[standard.lower()] = cls

    def get_notes(self, standard: str) -> Type | None:
        return self._notes.get(standard.lower())

    # ── Introspection ─────────────────────────────────────────────────────────

    def list_standards(self) -> list[str]:
        """Return all registered standards."""
        return sorted(self._standard_tables.keys())

    def summary(self) -> str:
        """Return a human-readable summary of registered plugins."""
        lines = ["Registered plugins:"]
        for std in self.list_standards():
            tables = self.get_all_table_types(std)
            val = "yes" if self.get_validator(std) else "no"
            alg = "yes" if self.get_aligner(std) else "no"
            nts = "yes" if self.get_notes(std) else "no"
            lines.append(
                f"  {std.upper()}: tables={tables}, "
                f"validator={val}, aligner={alg}, notes={nts}"
            )
        return "\n".join(lines)


# ── Global registry instance ──────────────────────────────────────────────────

registry = Registry()


def auto_discover() -> None:
    """Import all known plugin modules to trigger their self-registration.

    Call this once at startup (the runner does it automatically).
    Add new standard modules here as they are implemented.
    """
    _try_import("wmo.extractors.bufr")
    _try_import("wmo.validators.bufr")
    _try_import("wmo.aligners.reference")
    _try_import("wmo.notes.reference")
    _try_import("wmo.extractors.grib")
    _try_import("wmo.extractors.cct")


def _try_import(module_name: str) -> None:
    """Import a module, silently ignoring ImportError (missing dependency)."""
    try:
        __import__(module_name)
    except ImportError:
        pass
