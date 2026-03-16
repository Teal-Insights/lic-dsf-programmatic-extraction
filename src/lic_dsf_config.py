#!/usr/bin/env python3
"""
Shared configuration types and utilities for LIC-DSF programmatic extraction.

Template-specific configuration (workbook paths, export ranges, constraints,
region configs) lives in ``src/configs/<template>/config.py``. This module
provides shared type definitions and helper functions used across templates.
"""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Literal, TypedDict

import openpyxl.utils.cell
from excel_grapher import format_cell_key


class ExportRangeConfig(TypedDict):
    """
    Explicit range specification for export/annotation targets.

    Attributes:
        label: Human-readable label for the range (used for reporting only).
        range_spec: Sheet-qualified A1 range, e.g. "'Chart Data'!D10:D17".
        entrypoint_mode: Controls how export entrypoints are grouped for this
            range: "row_group" (one entrypoint per row) or "per_cell" (one
            entrypoint per cell, no row grouping).
    """

    label: str
    range_spec: str
    entrypoint_mode: Literal["row_group", "per_cell"]


def ensure_workbook_available(
    path: Path, url: str | None = None
) -> bool:
    """
    Ensure an LIC-DSF template workbook exists locally.

    If the workbook is missing, downloads it from ``url`` into ``path``.
    """
    if path.exists() and path.stat().st_size > 0:
        return True

    if url is None:
        return False

    from urllib.request import urlopen

    path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with urlopen(url, timeout=60) as resp:
            content_type = resp.headers.get("Content-Type", "")
            if "text/html" in content_type:
                raise ValueError(
                    f"URL returned HTML instead of a binary file "
                    f"(Content-Type: {content_type}). "
                    f"The download link may be broken: {url}"
                )

            with tempfile.NamedTemporaryFile(
                prefix=f".{path.name}.", suffix=".download", dir=str(path.parent), delete=False
            ) as tmp:
                shutil.copyfileobj(resp, tmp)
                tmp_path = Path(tmp.name)

        if tmp_path.stat().st_size == 0:
            tmp_path.unlink(missing_ok=True)
            return False

        tmp_path.replace(path)
        return True
    except Exception:
        return False


def parse_range_spec(spec: str) -> tuple[str, str]:
    """
    Parse a sheet-qualified range spec into (sheet_name, range_a1).

    Accepts specs like "'Chart Data'!D10:D17" or "Sheet1!A1:B2".
    """
    if "!" not in spec:
        raise ValueError(f"Range spec must contain '!': {spec!r}")
    sheet_part, range_part = spec.split("!", 1)
    sheet_part = sheet_part.strip()
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return sheet_part, range_part.strip()


def cells_in_range(sheet: str, range_a1: str) -> list[str]:
    """
    Expand an A1 range to a list of sheet-qualified cell keys.

    The returned keys use `excel_grapher.format_cell_key` so they match the
    dependency-graph expectations (including sheet-name quoting when needed).
    """
    if ":" in range_a1:
        start_a1, end_a1 = range_a1.split(":", 1)
        start_a1 = start_a1.strip()
        end_a1 = end_a1.strip()
    else:
        start_a1 = end_a1 = range_a1.strip()

    c1, r1 = openpyxl.utils.cell.coordinate_from_string(start_a1)
    c2, r2 = openpyxl.utils.cell.coordinate_from_string(end_a1)
    start_col_idx = openpyxl.utils.cell.column_index_from_string(c1)
    end_col_idx = openpyxl.utils.cell.column_index_from_string(c2)
    rlo, rhi = (r1, r2) if r1 <= r2 else (r2, r1)
    clo, chi = (start_col_idx, end_col_idx) if start_col_idx <= end_col_idx else (
        end_col_idx,
        start_col_idx,
    )

    out: list[str] = []
    for row in range(rlo, rhi + 1):
        for col_idx in range(clo, chi + 1):
            col_letter = openpyxl.utils.cell.get_column_letter(col_idx)
            out.append(format_cell_key(sheet, col_letter, row))
    return out


def discover_targets_from_ranges(
    export_ranges: list[ExportRangeConfig],
) -> list[str]:
    """
    Discover export/annotation targets from explicit range specs.

    Returns a de-duplicated list of sheet-qualified cell keys.
    """
    targets: list[str] = []
    for cfg in export_ranges:
        sheet_name, range_a1 = parse_range_spec(cfg["range_spec"])
        targets.extend(cells_in_range(sheet_name, range_a1))
    # Preserve order while de-duplicating.
    return list(dict.fromkeys(targets))
