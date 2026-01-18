#!/usr/bin/env python3
"""
Export LIC-DSF workbook formulas as standalone Python code.

This script discovers formula targets from `map_lic_dsf_indicators.INDICATOR_CONFIG`,
builds a dependency graph, and uses `excel-formula-expander`'s `CodeGenerator` to
emit a small Python package under `export/`.
"""

from __future__ import annotations

from pathlib import Path
import json
import re

import openpyxl
import openpyxl.utils.cell
from openpyxl.worksheet.worksheet import Worksheet
from excel_grapher import create_dependency_graph
from excel_formula_expander.codegen import CodeGenerator

from map_lic_dsf_indicators import (
    INDICATOR_CONFIG,
    WORKBOOK_PATH,
    discover_formula_cells_in_rows,
    ensure_workbook_available,
)


EXPORT_DIR = Path("export")
ENRICHMENT_AUDIT_PATH = Path("enrichment_audit.json")
MAX_DEPTH = 50


def discover_targets_by_indicator_row(workbook: Path) -> dict[tuple[str, int], list[str]]:
    targets_by_row: dict[tuple[str, int], list[str]] = {}
    for config in INDICATOR_CONFIG:
        sheet = config["sheet"]
        rows = config["indicator_rows"]
        for row in rows:
            targets = discover_formula_cells_in_rows(workbook, sheet, [row])
            if targets:
                targets_by_row[(sheet, row)] = list(dict.fromkeys(targets))
    return targets_by_row


def normalize_entrypoint_name(name: str) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z]+", "_", name.strip()).strip("_").lower()
    if not cleaned:
        return "sheet"
    if not cleaned[0].isalpha():
        return f"sheet_{cleaned}"
    return cleaned


def load_enrichment_audit_labels(path: Path) -> dict[tuple[str, int], list[str]]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}

    labels_by_row: dict[tuple[str, int], list[str]] = {}
    by_sheet = payload.get("by_sheet", {})
    for sheet, sheet_payload in by_sheet.items():
        for cell in sheet_payload.get("cells", []):
            address = cell.get("address")
            if not isinstance(address, str):
                continue
            match = re.match(r"^[A-Z]+(\d+)$", address)
            if not match:
                continue
            row = int(match.group(1))
            row_labels = cell.get("row_labels") or []
            if not isinstance(row_labels, list):
                continue
            if row_labels:
                labels_by_row.setdefault((sheet, row), []).extend(
                    label for label in row_labels if isinstance(label, str) and label.strip()
                )
    return labels_by_row


def build_entrypoints(
    targets_by_row: dict[tuple[str, int], list[str]]
) -> dict[str, list[str]]:
    labels_by_row = load_enrichment_audit_labels(ENRICHMENT_AUDIT_PATH)
    entrypoints: dict[str, list[str]] = {}
    for (sheet, row), targets in targets_by_row.items():
        label = next(iter(labels_by_row.get((sheet, row), [])), "")
        base = normalize_entrypoint_name(label or f"{sheet} {row}")
        name = base
        suffix = 2
        while name in entrypoints:
            name = f"{base}_{suffix}"
            suffix += 1
        entrypoints[name] = targets
    return entrypoints


def populate_leaf_values(graph, workbook: Path) -> None:
    """
    Populate values for leaf (non-formula) nodes from cached workbook values.

    Code generation only needs `node.value` for leaf cells; formulas are emitted
    from `node.formula`.
    """
    wb = openpyxl.load_workbook(workbook, data_only=True, keep_vba=True)
    try:
        worksheets: dict[str, Worksheet] = {}
        for key in graph:
            node = graph.get_node(key)
            if node is None or node.formula is not None:
                continue
            if node.sheet not in worksheets:
                if node.sheet not in wb.sheetnames:
                    continue
                worksheets[node.sheet] = wb[node.sheet]
            ws = worksheets[node.sheet]
            col_idx = openpyxl.utils.cell.column_index_from_string(node.column)
            node.value = ws.cell(row=node.row, column=col_idx).value
    finally:
        wb.close()


def main() -> None:
    if not ensure_workbook_available(WORKBOOK_PATH):
        print(f"Error: Workbook not available at {WORKBOOK_PATH}")
        return

    targets_by_row = discover_targets_by_indicator_row(WORKBOOK_PATH)
    targets: list[str] = []
    for row_targets in targets_by_row.values():
        targets.extend(row_targets)
    targets = list(dict.fromkeys(targets))
    if not targets:
        print("No targets found. Nothing to export.")
        return

    print("=" * 70)
    print("LIC-DSF Workbook Export (standalone Python code)")
    print("=" * 70)
    print(f"Workbook: {WORKBOOK_PATH}")
    entrypoints = build_entrypoints(targets_by_row)
    print(f"Targets:  {len(targets)}")
    print(f"Entrypoints: {len(entrypoints)}")

    print("\nBuilding dependency graph...")
    graph = create_dependency_graph(
        WORKBOOK_PATH,
        targets,
        load_values=False,
        max_depth=MAX_DEPTH,
    )

    print("Populating leaf values from cached workbook values...")
    populate_leaf_values(graph, WORKBOOK_PATH)

    print("Generating Python package...")
    generator = CodeGenerator(graph)
    modules = generator.generate_modules_in_package(
        targets,
        package_name=WORKBOOK_PATH.stem,
        entrypoints=entrypoints if entrypoints else None,
    )

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    for rel, content in modules.items():
        dst = EXPORT_DIR / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(content, encoding="utf-8")

    pkg_prefix = Path(next(iter(modules.keys()))).parts[0] if modules else "(unknown)"
    print(f"\nWrote {len(modules)} files under {EXPORT_DIR / pkg_prefix}")


if __name__ == "__main__":
    main()

