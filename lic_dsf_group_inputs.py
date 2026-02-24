#!/usr/bin/env python3
"""
Group hardcoded (non-formula) input cells into semantically labeled clusters.

This script:
- Discovers formula targets in configured indicator rows (see `INDICATOR_CONFIG`)
- Builds a dependency graph from those targets
- Extracts leaf, non-formula cells (the same concept as `export/lic_dsf/inputs.py`)
- Enriches those cells with row/column labels
- Groups inputs by (sheet, labels), ignoring any axis that contains year labels

The output is a JSON file that can be used to generate higher-level input setters
(e.g., row-wise time series setters and small-table setters).
"""

from __future__ import annotations

import argparse
import ast
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Literal

import openpyxl.utils.cell
from excel_grapher import DependencyGraph

from lic_dsf_labels import WORKBOOK_PATH, ensure_workbook_available, is_offset_label
from lic_dsf_pipeline import (
    BLANK_CONSTANT_EXCLUDES,
    STRING_CONSTANT_EXCLUDES,
    build_graph,
    classify_input_addresses,
    discover_targets,
    enrich_graph,
    iter_string_constant_addresses,
    populate_leaf_values,
)

_SAFE_SHEET_NAME_RE = re.compile(r"^[A-Za-z_][0-9A-Za-z_]*$")


def _format_sheet_name(sheet: str) -> str:
    """
    Format a sheet name for sheet-qualified Excel addresses.

    Only quote sheet names when needed (e.g., spaces or punctuation), and escape
    embedded single quotes by doubling them.
    """
    if _SAFE_SHEET_NAME_RE.match(sheet):
        return sheet
    escaped = sheet.replace("'", "''")
    return f"'{escaped}'"


def _format_address(sheet: str, a1: str) -> str:
    return f"{_format_sheet_name(sheet)}!{a1}"


def _normalize_label(label: str) -> str:
    return " ".join(label.split()).strip()


def _is_year_label(label: str) -> bool:
    text = _normalize_label(label)
    if is_offset_label(text):
        return True
    try:
        year = int(text)
    except ValueError:
        return False
    return 1900 <= year <= 2100


def _contains_year(labels: Iterable[str]) -> bool:
    return any(_is_year_label(label) for label in labels)


def _labels_key(labels: Iterable[str]) -> tuple[str, ...]:
    normalized = [_normalize_label(label) for label in labels if isinstance(label, str)]
    cleaned = [label for label in normalized if label]
    # Sort + de-dupe for stable set-like behavior.
    return tuple(sorted(dict.fromkeys(cleaned)))


def _non_year_labels(labels: Iterable[str]) -> list[str]:
    return [
        label
        for label in labels
        if isinstance(label, str) and _normalize_label(label) and not _is_year_label(label)
    ]


GroupMode = Literal["ignore_column_axis_years", "ignore_row_axis_years", "no_year_axis", "both_axes_years"]


def _group_mode(row_labels: list[str], column_labels: list[str]) -> GroupMode:
    row_has_year = _contains_year(row_labels)
    col_has_year = _contains_year(column_labels)
    if col_has_year and not row_has_year:
        return "ignore_column_axis_years"
    if row_has_year and not col_has_year:
        return "ignore_row_axis_years"
    if not row_has_year and not col_has_year:
        return "no_year_axis"
    return "both_axes_years"


@dataclass(frozen=True, slots=True)
class GroupKey:
    sheet: str
    mode: GroupMode
    row_labels_key: tuple[str, ...]
    column_labels_key: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class InputCell:
    address: str  # sheet-qualified (with quoting if needed)
    sheet: str
    row: int
    col: int
    col_letter: str
    row_labels: list[str]
    column_labels: list[str]


def _load_export_default_input_addresses(path: Path) -> set[str]:
    """
    Load `DEFAULT_INPUTS` keys from a generated `export/.../inputs.py` module.

    This is useful when you want grouping to match the currently-generated
    runtime package exactly.
    """
    src = path.read_text(encoding="utf-8")
    mod = ast.parse(src)
    for node in mod.body:
        if not isinstance(node, ast.Assign):
            continue
        if not any(getattr(t, "id", None) == "DEFAULT_INPUTS" for t in node.targets):
            continue
        default_inputs = ast.literal_eval(node.value)
        if not isinstance(default_inputs, dict):
            raise ValueError("DEFAULT_INPUTS is not a dict")
        return {str(k) for k in default_inputs.keys()}
    raise ValueError("DEFAULT_INPUTS not found")


def iter_input_cells(
    graph: DependencyGraph,
    enrichment_results: dict[str, dict[str, Any]],
) -> list[InputCell]:
    inputs: list[InputCell] = []
    for node_key, data in enrichment_results.items():
        node = graph.get_node(node_key)
        if node is None:
            continue
        if node.formula is not None:
            continue
        if not node.is_leaf:
            continue

        row_labels = data.get("row_labels") or []
        column_labels = data.get("column_labels") or []
        if not isinstance(row_labels, list) or not isinstance(column_labels, list):
            continue

        col_idx = openpyxl.utils.cell.column_index_from_string(node.column)
        address = _format_address(node.sheet, node.address)

        inputs.append(
            InputCell(
                address=address,
                sheet=node.sheet,
                row=node.row,
                col=col_idx,
                col_letter=node.column,
                row_labels=[str(label) for label in row_labels if isinstance(label, str)],
                column_labels=[str(label) for label in column_labels if isinstance(label, str)],
            )
        )
    return inputs


def build_input_groups_payload(
    *,
    targets: list[str],
    graph: DependencyGraph,
    input_cells: list[InputCell],
    restricted_to_export_default_inputs: bool = False,
    export_default_inputs_count: int | None = None,
) -> dict[str, Any]:
    groups: dict[GroupKey, list[InputCell]] = {}
    for cell in input_cells:
        key = _key_for_cell(cell)
        groups.setdefault(key, []).append(cell)

    keys_sorted = sorted(
        groups.keys(),
        key=lambda k: (k.sheet, k.mode, k.row_labels_key, k.column_labels_key),
    )

    payload_groups: list[dict[str, Any]] = []
    for idx, key in enumerate(keys_sorted, start=1):
        cells = groups[key]
        cells_sorted = sorted(cells, key=lambda c: (c.row, c.col))
        rect = _rectangular_range(cells_sorted)

        group_payload: dict[str, Any] = {
            "group_id": f"g{idx:05d}",
            "sheet": key.sheet,
            "mode": key.mode,
            "row_labels_key": list(key.row_labels_key),
            "column_labels_key": list(key.column_labels_key),
            "row_labels_has_year": _contains_year(cells_sorted[0].row_labels)
            if cells_sorted
            else False,
            "column_labels_has_year": _contains_year(cells_sorted[0].column_labels)
            if cells_sorted
            else False,
            "example_row_labels": cells_sorted[0].row_labels if cells_sorted else [],
            "example_column_labels": cells_sorted[0].column_labels if cells_sorted else [],
            "cell_count": len(cells_sorted),
            "cells": [c.address for c in cells_sorted],
        }

        if rect is not None:
            bbox, range_a1, shape = rect
            group_payload["bounding_box"] = bbox
            group_payload["range_a1"] = range_a1
            group_payload["shape"] = {"rows": shape[0], "cols": shape[1]}
        else:
            group_payload["bounding_box"] = None
            group_payload["range_a1"] = None
            group_payload["shape"] = None

        payload_groups.append(group_payload)

    summary = {
        "targets": len(targets),
        "graph_nodes": len(graph),
        "input_cells": len(input_cells),
        "groups": len(payload_groups),
        "restricted_to_export_default_inputs": bool(restricted_to_export_default_inputs),
        "export_default_inputs_count": export_default_inputs_count,
        "groups_by_mode": {
            mode: sum(1 for g in payload_groups if g["mode"] == mode)
            for mode in (
                "ignore_column_axis_years",
                "ignore_row_axis_years",
                "no_year_axis",
                "both_axes_years",
            )
        },
    }

    return {
        "workbook": str(WORKBOOK_PATH),
        "summary": summary,
        "groups": payload_groups,
    }


def _key_for_cell(cell: InputCell) -> GroupKey:
    mode = _group_mode(cell.row_labels, cell.column_labels)

    if mode == "ignore_column_axis_years":
        return GroupKey(
            sheet=cell.sheet,
            mode=mode,
            row_labels_key=_labels_key(cell.row_labels),
            column_labels_key=(),
        )
    if mode == "ignore_row_axis_years":
        return GroupKey(
            sheet=cell.sheet,
            mode=mode,
            row_labels_key=_labels_key(_non_year_labels(cell.row_labels)),
            column_labels_key=_labels_key(cell.column_labels),
        )
    if mode == "no_year_axis":
        return GroupKey(
            sheet=cell.sheet,
            mode=mode,
            row_labels_key=_labels_key(cell.row_labels),
            column_labels_key=_labels_key(cell.column_labels),
        )
    # both_axes_years: fall back to non-year labels on both axes.
    return GroupKey(
        sheet=cell.sheet,
        mode=mode,
        row_labels_key=_labels_key(_non_year_labels(cell.row_labels)),
        column_labels_key=_labels_key(_non_year_labels(cell.column_labels)),
    )


def _rectangular_range(cells: list[InputCell]) -> tuple[dict[str, int], str, tuple[int, int]] | None:
    if not cells:
        return None
    if len({c.sheet for c in cells}) != 1:
        return None

    min_row = min(c.row for c in cells)
    max_row = max(c.row for c in cells)
    min_col = min(c.col for c in cells)
    max_col = max(c.col for c in cells)
    expected = (max_row - min_row + 1) * (max_col - min_col + 1)
    if expected != len(cells):
        return None

    present = {(c.row, c.col) for c in cells}
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            if (r, c) not in present:
                return None

    sheet = cells[0].sheet
    start = f"{openpyxl.utils.cell.get_column_letter(min_col)}{min_row}"
    end = f"{openpyxl.utils.cell.get_column_letter(max_col)}{max_row}"
    range_a1 = f"{_format_sheet_name(sheet)}!{start}:{end}"
    shape = (max_row - min_row + 1, max_col - min_col + 1)
    bbox = {
        "start_row": min_row,
        "end_row": max_row,
        "start_col": min_col,
        "end_col": max_col,
    }
    return bbox, range_a1, shape


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Group LIC-DSF template hardcoded input cells into labeled clusters."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=WORKBOOK_PATH,
        help="Path to workbook (default: workbooks/lic-dsf-template-2026-01-31.xlsm)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("input_groups.json"),
        help="Output JSON path (default: input_groups.json)",
    )
    parser.add_argument("--max-depth", type=int, default=50, help="Dependency graph max depth")
    parser.add_argument(
        "--restrict-to-export-default-inputs",
        action="store_true",
        help="If set, only include inputs present in --export-inputs-path DEFAULT_INPUTS.",
    )
    parser.add_argument(
        "--export-inputs-path",
        type=Path,
        default=Path("export/lic_dsf/inputs.py"),
        help="Path to generated inputs module (default: export/lic_dsf/inputs.py)",
    )
    args = parser.parse_args()

    workbook: Path = args.workbook
    if workbook == WORKBOOK_PATH and not ensure_workbook_available(WORKBOOK_PATH):
        raise SystemExit(f"Workbook not available at {WORKBOOK_PATH}")
    if not workbook.exists():
        raise SystemExit(f"Workbook not found: {workbook}")

    print("=" * 70)
    print("LIC-DSF Input Grouping")
    print("=" * 70)
    print(f"Workbook: {workbook}")

    print("\n1. Discovering formula targets in indicator rows...")
    targets = discover_targets(workbook)
    print(f"   Targets: {len(targets)}")
    if not targets:
        raise SystemExit("No targets found. Nothing to group.")

    print("\n2. Building dependency graph...")
    graph = build_graph(workbook, targets, max_depth=args.max_depth)
    print(f"   Graph nodes: {len(graph)}")

    print("\n2b. Populating leaf values from cached workbook values...")
    populate_leaf_values(graph, workbook)

    print("\n3. Enriching nodes with row/column labels...")
    enrichment_results = enrich_graph(graph, workbook)
    print(f"   Enriched nodes: {len(enrichment_results)}")

    print("\n4. Extracting leaf input cells...")
    input_cells = iter_input_cells(graph, enrichment_results)
    print(f"   Input cells (leaf, non-formula): {len(input_cells)}")

    print("\n4b. Classifying constants vs inputs...")
    constant_ranges = iter_string_constant_addresses(graph, STRING_CONSTANT_EXCLUDES)
    input_addresses = classify_input_addresses(
        graph,
        targets,
        constant_ranges=constant_ranges,
        constant_blanks=True,
        blank_excludes=BLANK_CONSTANT_EXCLUDES,
    )
    before = len(input_cells)
    input_cells = [c for c in input_cells if c.address in input_addresses]
    print(f"   Inputs after constant filtering: {len(input_cells)} (dropped {before - len(input_cells)})")

    export_default_inputs: set[str] | None = None
    if args.restrict_to_export_default_inputs:
        export_default_inputs = _load_export_default_input_addresses(args.export_inputs_path)
        before = len(input_cells)
        input_cells = [c for c in input_cells if c.address in export_default_inputs]
        print(
            f"   Restricted to export DEFAULT_INPUTS: {len(input_cells)} "
            f"(dropped {before - len(input_cells)})"
        )

    print("\n5. Grouping inputs by labels (ignoring year axes)...")
    output_payload = build_input_groups_payload(
        targets=targets,
        graph=graph,
        input_cells=input_cells,
        restricted_to_export_default_inputs=bool(args.restrict_to_export_default_inputs),
        export_default_inputs_count=len(export_default_inputs)
        if export_default_inputs is not None
        else None,
    )

    args.output.write_text(
        json.dumps(output_payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print("\n6. Wrote grouped inputs:")
    print(f"   {args.output}")
    print(f"   Groups: {len(payload_groups)}")


if __name__ == "__main__":
    main()

