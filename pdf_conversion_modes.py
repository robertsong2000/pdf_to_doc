#!/usr/bin/env python3
"""
PDF to DOCX conversion mode settings.

The default mode preserves the existing pdf2docx behavior. Alternative modes
are explicit user choices because they can trade layout fidelity for text flow.
"""

CONVERSION_MODE_DEFAULT = "layout"
CONVERSION_MODE_TEXT = "text"
CONVERSION_MODE_NO_LATTICE = "no_lattice"

CONVERSION_MODE_OPTIONS = {
    CONVERSION_MODE_DEFAULT: {},
    CONVERSION_MODE_TEXT: {
        "parse_lattice_table": False,
        "parse_stream_table": False,
    },
    CONVERSION_MODE_NO_LATTICE: {
        "parse_lattice_table": False,
    },
}


def get_conversion_options(conversion_mode: str | None) -> dict:
    mode = conversion_mode or CONVERSION_MODE_DEFAULT
    if mode not in CONVERSION_MODE_OPTIONS:
        valid_modes = ", ".join(sorted(CONVERSION_MODE_OPTIONS))
        raise ValueError(f"Unsupported conversion mode: {mode}. Valid modes: {valid_modes}")
    return dict(CONVERSION_MODE_OPTIONS[mode])
