#!/usr/bin/env python3
"""
Export LIC-DSF workbook formulas as standalone Python code.

This script discovers formula targets from `map_lic_dsf_indicators.INDICATOR_CONFIG`,
builds a dependency graph, and uses `excel-formula-expander`'s `CodeGenerator` to
emit a small Python package under `export/`.
"""

from __future__ import annotations

from pathlib import Path

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
MAX_DEPTH = 50


def discover_targets(workbook: Path) -> list[str]:
    all_targets: list[str] = []
    for config in INDICATOR_CONFIG:
        sheet = config["sheet"]
        rows = config["indicator_rows"]
        all_targets.extend(discover_formula_cells_in_rows(workbook, sheet, rows))
    # Preserve order while de-duplicating.
    return list(dict.fromkeys(all_targets))


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

    targets = discover_targets(WORKBOOK_PATH)
    if not targets:
        print("No targets found. Nothing to export.")
        return

    print("=" * 70)
    print("LIC-DSF Workbook Export (standalone Python code)")
    print("=" * 70)
    print(f"Workbook: {WORKBOOK_PATH}")
    print(f"Targets:  {len(targets)}")

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
    modules = generator.generate_modules_in_package(targets, package_name=WORKBOOK_PATH.stem)

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    for rel, content in modules.items():
        dst = EXPORT_DIR / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(content, encoding="utf-8")

    pkg_prefix = Path(next(iter(modules.keys()))).parts[0] if modules else "(unknown)"
    print(f"\nWrote {len(modules)} files under {EXPORT_DIR / pkg_prefix}")


if __name__ == "__main__":
    main()

