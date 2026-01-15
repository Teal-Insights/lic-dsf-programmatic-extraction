#!/usr/bin/env python3
"""
Export a slice of the LIC-DSF workbook as standalone Python code.

This script discovers formula targets (by default, the configured indicator rows),
builds a dependency graph including cached cell values, and uses the
excel-formula-expander CodeGenerator to emit standalone Python code that can
compute those targets.

The generated module is written under `export/` by default.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl
import openpyxl.utils.cell
from excel_grapher import create_dependency_graph
from formula_expander.codegen import CodeGenerator

from map_lic_dsf_indicators import (
    INDICATOR_CONFIG,
    WORKBOOK_PATH,
    discover_formula_cells_in_rows,
    ensure_workbook_available,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export workbook formulas to standalone Python code."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=None,
        help=(
            "Path to the LIC-DSF workbook (.xlsm). "
            "If omitted, uses the default template workbook."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=(
            "Output .py path. If omitted, writes to export/<workbook-stem>.py "
            "relative to the repository root."
        ),
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=50,
        help="Maximum dependency depth to traverse when building the graph.",
    )
    parser.add_argument(
        "--max-targets",
        type=int,
        default=None,
        help=(
            "If set, cap the number of discovered targets (useful for quick iteration). "
            "Ignored when --target is provided."
        ),
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=0,
        help=(
            "If > 0, export multiple modules with at most this many targets each "
            "(export/<stem>_part001.py, export/<stem>_part002.py, ...)."
        ),
    )
    parser.add_argument(
        "--embed-values",
        action=argparse.BooleanOptionalAction,
        default=True,
        help=(
            "If enabled, populate leaf cell values from the workbook's cached values. "
            "Disable to keep all leaf values as 0/empty in the generated module."
        ),
    )
    parser.add_argument(
        "--target",
        dest="targets",
        action="append",
        default=None,
        help=(
            "Sheet-qualified target cell address (e.g., 'Sheet1!B10'). "
            "Repeatable. If omitted, targets are auto-discovered from INDICATOR_CONFIG."
        ),
    )
    parser.add_argument(
        "--preview-lines",
        type=int,
        default=0,
        help="If > 0, print the first N lines of the generated code to stdout.",
    )
    return parser.parse_args()


def resolve_workbook_path(arg: Path | None) -> Path | None:
    if arg is not None:
        if not (arg.exists() and arg.stat().st_size > 0):
            return None
        return arg

    if not ensure_workbook_available(WORKBOOK_PATH):
        return None
    return WORKBOOK_PATH


def discover_default_targets(workbook: Path) -> list[str]:
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
        worksheets: dict[str, object] = {}
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
    args = parse_args()

    workbook = resolve_workbook_path(args.workbook)
    if workbook is None:
        if args.workbook is None:
            print(f"Error: Workbook not available at {WORKBOOK_PATH}")
        else:
            print(f"Error: Workbook not available at {args.workbook}")
        return

    targets = args.targets if args.targets else discover_default_targets(workbook)
    if args.targets is None and args.max_targets is not None:
        targets = targets[: args.max_targets]

    targets_count = len(targets)
    if targets_count == 0:
        print("No targets found. Nothing to export.")
        return

    out_path = args.output
    if out_path is None:
        out_path = Path("export") / f"{workbook.stem}.py"

    out_path.parent.mkdir(parents=True, exist_ok=True)

    print("=" * 70, flush=True)
    print("LIC-DSF Workbook Export (standalone Python code)", flush=True)
    print("=" * 70, flush=True)
    print(f"Workbook: {workbook}", flush=True)
    print(f"Targets:  {targets_count}", flush=True)
    print(f"Output:   {out_path}", flush=True)

    chunk_size = args.chunk_size if args.chunk_size and args.chunk_size > 0 else 0
    if chunk_size == 0:
        print(
            f"\nBuilding dependency graph for {len(targets)} targets...",
            flush=True,
        )
        graph = create_dependency_graph(
            workbook,
            targets,
            load_values=False,
            max_depth=args.max_depth,
        )

        if args.embed_values:
            print("Populating leaf values...", flush=True)
            populate_leaf_values(graph, workbook)

        print("Generating Python code...", flush=True)
        code = CodeGenerator(graph).generate(targets)
        out_path.write_text(code, encoding="utf-8")
        print(f"Wrote {out_path}", flush=True)

        if args.preview_lines and args.preview_lines > 0:
            preview = "\n".join(code.splitlines()[: args.preview_lines])
            print("\n--- Preview ---")
            print(preview)
        print("\nDone.", flush=True)
        return

    chunks: list[list[str]] = [
        targets[i : i + chunk_size] for i in range(0, len(targets), chunk_size)
    ]

    for i, chunk_targets in enumerate(chunks, start=1):
        chunk_out = out_path.with_name(f"{out_path.stem}_part{i:03d}{out_path.suffix}")

        print(
            f"\n[{i}/{len(chunks)}] Building dependency graph for {len(chunk_targets)} targets...",
            flush=True,
        )
        graph = create_dependency_graph(
            workbook,
            chunk_targets,
            load_values=False,
            max_depth=args.max_depth,
        )

        if args.embed_values:
            print(f"[{i}/{len(chunks)}] Populating leaf values...", flush=True)
            populate_leaf_values(graph, workbook)

        print(f"[{i}/{len(chunks)}] Generating Python code...", flush=True)
        code = CodeGenerator(graph).generate(chunk_targets)
        chunk_out.write_text(code, encoding="utf-8")
        print(f"[{i}/{len(chunks)}] Wrote {chunk_out}", flush=True)

        if args.preview_lines and args.preview_lines > 0:
            preview = "\n".join(code.splitlines()[: args.preview_lines])
            print("\n--- Preview ---")
            print(preview)
            args.preview_lines = 0  # Only preview the first generated module.

    print("\nDone.", flush=True)


if __name__ == "__main__":
    main()

