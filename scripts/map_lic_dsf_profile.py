#!/usr/bin/env python3
"""
Profiling variant of map_lic_dsf.py for cProfile (often with `timeout`).

Same export targets and graph workflow as map_lic_dsf.py, but:
- No LicDsfConstraints entries (empty TypedDict).
- USE_CACHED_DYNAMIC_REFS = True so OFFSET/INDIRECT resolve from cached
  workbook values and DynamicRefConfig is not built from constraints.

`create_dependency_graph` always calls `fastpyxl.load_workbook` (formula book plus a
lazy `data_only=True` load for cached dynamic refs when `load_values=False`).
That dominates wall time, so a naive cProfile run mostly measures XML parsing, not
the dependency walk.

Use ``--profile-post-load`` to pre-load both workbooks, monkeypatch
``fastpyxl.load_workbook`` to return those instances, then profile only
``create_dependency_graph`` so hotspots reflect the post-load graph phase.

Examples:

    uv run python scripts/map_lic_dsf_profile.py --profile-post-load map_post_load.prof
    timeout 600 uv run python scripts/map_lic_dsf_profile.py \\
        --profile-post-load map_post_load.prof --profile-dump-on-sigterm
    uv run python -m pstats map_post_load.prof
"""

from __future__ import annotations

import argparse
import cProfile
import logging
import signal
import sys
import time
from collections.abc import Callable
from pathlib import Path
from typing import Literal, TypedDict

import fastpyxl
import fastpyxl.utils.cell

from excel_grapher import (
    CycleError,
    DependencyGraph,
    DynamicRefConfig,
    DynamicRefError,
    create_dependency_graph,
    format_cell_key,
    get_calc_settings,
    to_graphviz,
    validate_graph,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)s %(message)s",
    datefmt="%H:%M:%S",
    force=True,
)
logging.getLogger("excel_grapher.grapher.dynamic_refs").setLevel(logging.INFO)


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

# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

WORKBOOK_PATH = Path("workbooks/lic-dsf-template-2025-08-12.xlsm")
WORKBOOK_TEMPLATE_URL = (
    "https://thedocs.worldbank.org/en/doc/f0ade6bcf85b6f98dbeb2c39a2b7770c-0360012025/original/LIC-DSF-IDA21-Template-08-12-2025-vf.xlsm"
)

USE_CACHED_DYNAMIC_REFS = True

# ---------------------------------------------------------------------------
# Export package
# ---------------------------------------------------------------------------

PACKAGE_NAME = "lic_dsf_2025_08_12"
EXPORT_DIR = Path("dist/lic-dsf-2025-08-12")

# ---------------------------------------------------------------------------
# Export ranges
# ---------------------------------------------------------------------------

STRESS_TEST_ROW_LABELS: list[str] = [
    "Baseline",
    "A1. Key variables at their historical averages in 2024-2034 2/",
    "B1. Real GDP growth",
    "B2. Primary balance",
    "B3. Exports",
    "B4. Other flows 3/",
    "B5. Depreciation",
    "B6. Combination of B1-B5",
    "",
    "C1. Combined contingent liabilities",
    "C2. Natural disaster",
    "C3. Commodity price",
    "C4. Market Financing",
    "A2. Alternative Scenario :[Customize, enter title]",
]


STRESS_TEST_BLOCKS: list[tuple[str, int]] = [
    ("PV of Debt-to-GDP Ratio", 239),
    ("PV of Debt-to-Revenue Ratio", 281),
    ("Debt Service-to-Revenue Ratio", 318),
    ("Debt Service-to-GDP Ratio", 351),
]


FIGURE_DATA_ROWS: list[int] = [
    # Figure 1 (Output 2-1 Stress_Charts_Ex)
    51,
    61,
    62,
    63,
    64,
    66,
    93,
    103,
    104,
    105,
    106,
    108,
    135,
    145,
    146,
    147,
    148,
    150,
    177,
    187,
    188,
    189,
    190,
    192,
    # Figure 2 extras (Output 2-2 Stress_Charts_Pub)
    263,
    264,
    265,
    267,
    306,
    341,
    342,
    343,
]


EXPORT_FIXED_RANGES: list[ExportRangeConfig] = [
    {
        "label": "External DSA risk rating signals",
        "range_spec": "'Chart Data'!D10:D17",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Fiscal (Total Public Debt) risk rating signals",
        "range_spec": "'Chart Data'!I10:I14",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Applicable tailored stress test signals",
        "range_spec": "'Chart Data'!I17:I19",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Fiscal space for moderate risk category",
        "range_spec": "'Chart Data'!E25:E27",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Overall rating",
        "range_spec": "'Chart Data'!L10:L11",
        "entrypoint_mode": "row_group",
    },
]


def _export_chart_data_ranges() -> list[ExportRangeConfig]:
    out: list[ExportRangeConfig] = list(EXPORT_FIXED_RANGES)
    seen_row_specs = {entry["range_spec"] for entry in out}

    def add_chart_data_row(row: int, label: str) -> None:
        range_spec = f"'Chart Data'!D{row}:X{row}"
        if range_spec in seen_row_specs:
            return
        out.append(
            {
                "label": label,
                "range_spec": range_spec,
                "entrypoint_mode": "row_group",
            }
        )
        seen_row_specs.add(range_spec)

    for metric_label, start_row in STRESS_TEST_BLOCKS:
        for i, row_label in enumerate(STRESS_TEST_ROW_LABELS):
            if not row_label:
                continue
            row = start_row + i
            add_chart_data_row(row, f"{metric_label} - {row_label}")

    for row in FIGURE_DATA_ROWS:
        add_chart_data_row(row, f"Figure data row {row}")

    return out


EXPORT_RANGES: list[ExportRangeConfig] = _export_chart_data_ranges()


def _patch_fastpyxl_workbook_cache(workbook_path: Path) -> Callable[[], None]:
    """
    Load formula and values workbooks once, then patch fastpyxl.load_workbook so
    create_dependency_graph's internal loads are cheap dictionary lookups.

    excel_grapher passes only ``data_only`` and ``keep_vba`` besides the path;
    other kwargs match fastpyxl defaults.
    """
    resolved = workbook_path.resolve()
    keep_vba = workbook_path.suffix.lower() == ".xlsm"
    print(
        f"   Pre-loading workbook (formulas + values) at {workbook_path}...",
        flush=True,
    )
    t0 = time.perf_counter()
    wb_formulas = fastpyxl.load_workbook(
        workbook_path,
        read_only=False,
        keep_vba=keep_vba,
        data_only=False,
        keep_links=True,
        rich_text=False,
    )
    wb_values = fastpyxl.load_workbook(
        workbook_path,
        read_only=False,
        keep_vba=keep_vba,
        data_only=True,
        keep_links=True,
        rich_text=False,
    )
    print(f"   Pre-load finished in {time.perf_counter() - t0:.2f}s", flush=True)

    original = fastpyxl.load_workbook

    def patched(
        filename: str | Path,
        read_only: bool = False,
        keep_vba: bool = True,
        data_only: bool = False,
        keep_links: bool = True,
        rich_text: bool = False,
    ) -> fastpyxl.Workbook:
        try:
            if Path(filename).resolve() != resolved:
                return original(
                    filename,
                    read_only=read_only,
                    keep_vba=keep_vba,
                    data_only=data_only,
                    keep_links=keep_links,
                    rich_text=rich_text,
                )
        except (OSError, TypeError, ValueError):
            return original(
                filename,
                read_only=read_only,
                keep_vba=keep_vba,
                data_only=data_only,
                keep_links=keep_links,
                rich_text=rich_text,
            )
        return wb_values if data_only else wb_formulas

    fastpyxl.load_workbook = patched  # type: ignore[method-assign]

    def restore() -> None:
        fastpyxl.load_workbook = original  # type: ignore[method-assign]

    return restore


# ---------------------------------------------------------------------------
# Constraints (intentionally empty — use scripts/map_lic_dsf.py for full set)
# ---------------------------------------------------------------------------


class LicDsfConstraints(TypedDict, total=False):
    pass


def get_dynamic_ref_config() -> DynamicRefConfig:
    """Unused when USE_CACHED_DYNAMIC_REFS is True; mirrors map_lic_dsf API."""
    return DynamicRefConfig.from_constraints_and_workbook(
        LicDsfConstraints,
        WORKBOOK_PATH,
    )


def create_graph_for_targets(all_targets: list[str]) -> DependencyGraph:
    """Run create_dependency_graph with the same flags as main() step 2."""
    dynamic_refs: DynamicRefConfig | None = None
    if not USE_CACHED_DYNAMIC_REFS:
        dynamic_refs = DynamicRefConfig.from_constraints_and_workbook(
            LicDsfConstraints,
            WORKBOOK_PATH,
        )
    return create_dependency_graph(
        WORKBOOK_PATH,
        all_targets,
        load_values=False,
        max_depth=50,
        dynamic_refs=dynamic_refs,
        use_cached_dynamic_refs=USE_CACHED_DYNAMIC_REFS,
    )


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

    range_a1 may be a single cell ("D10") or a range ("D10:D17", "D239:X252").
    """
    if ":" in range_a1:
        start_a1, end_a1 = range_a1.split(":", 1)
        start_a1 = start_a1.strip()
        end_a1 = end_a1.strip()
    else:
        start_a1 = end_a1 = range_a1.strip()

    c1, r1 = fastpyxl.utils.cell.coordinate_from_string(start_a1)
    c2, r2 = fastpyxl.utils.cell.coordinate_from_string(end_a1)
    start_col_idx = fastpyxl.utils.cell.column_index_from_string(c1)
    end_col_idx = fastpyxl.utils.cell.column_index_from_string(c2)
    rlo, rhi = (r1, r2) if r1 <= r2 else (r2, r1)
    clo, chi = (start_col_idx, end_col_idx) if start_col_idx <= end_col_idx else (end_col_idx, start_col_idx)

    out: list[str] = []
    for row in range(rlo, rhi + 1):
        for col_idx in range(clo, chi + 1):
            col_letter = fastpyxl.utils.cell.get_column_letter(col_idx)
            out.append(format_cell_key(sheet, col_letter, row))
    return out


def collect_export_targets() -> list[str]:
    """Sheet-qualified cell keys for all export ranges (same set as main())."""
    all_targets: list[str] = []
    for entry in EXPORT_RANGES:
        spec = entry["range_spec"]
        sheet_name, range_a1 = parse_range_spec(spec)
        all_targets.extend(cells_in_range(sheet_name, range_a1))
    return all_targets


def profile_post_load_graph(
    *,
    output_prof: Path,
    dump_on_sigterm: bool,
) -> None:
    """
    Profile only ``create_dependency_graph`` with workbook I/O moved out of the profile.

    Pre-loads formula and values workbooks, patches ``fastpyxl.load_workbook``, then
    runs the profiler around ``create_graph_for_targets``.
    """
    if not WORKBOOK_PATH.exists():
        print(f"Error: Workbook not found at {WORKBOOK_PATH}")
        sys.exit(1)

    all_targets = collect_export_targets()
    print(f"   Export targets: {len(all_targets)} cells", flush=True)

    restore_load_workbook = _patch_fastpyxl_workbook_cache(WORKBOOK_PATH)

    prof = cProfile.Profile()
    dumped = False

    def dump_prof() -> None:
        nonlocal dumped
        if dumped:
            return
        prof.disable()
        output_prof.parent.mkdir(parents=True, exist_ok=True)
        prof.dump_stats(str(output_prof))
        dumped = True
        print(f"   Wrote {output_prof}", flush=True)
        print("\nTop 40 by cumulative time (strip_dirs):", flush=True)
        import pstats

        pstats.Stats(str(output_prof)).strip_dirs().sort_stats("cumulative").print_stats(40)

    def on_sigterm(_signum: int, _frame: object | None) -> None:
        dump_prof()
        sys.exit(124)

    if dump_on_sigterm:
        signal.signal(signal.SIGTERM, on_sigterm)

    print("\nProfiling create_dependency_graph (post-load / patched load_workbook)...", flush=True)
    t0 = time.perf_counter()
    try:
        prof.enable()
        create_graph_for_targets(all_targets)
    finally:
        dump_prof()
        restore_load_workbook()

    print(f"   Profiled graph build wall time: {time.perf_counter() - t0:.2f}s", flush=True)


def main() -> None:
    print("=" * 70)
    print("LIC-DSF Indicator Dependency Mapping (profile build)")
    print("=" * 70)

    if not WORKBOOK_PATH.exists():
        print(f"Error: Workbook not found at {WORKBOOK_PATH}")
        return

    # Discover targets: explicit ranges (all cells) and indicator rows (formula cells only)
    print("\n1. Collecting target cells...")
    all_targets: list[str] = []

    for entry in EXPORT_RANGES:
        label = entry["label"]
        spec = entry["range_spec"]
        sheet_name, range_a1 = parse_range_spec(spec)
        targets = cells_in_range(sheet_name, range_a1)
        print(f"   {label}: {spec} -> {len(targets)} cells")
        all_targets.extend(targets)

    print(f"\n   Total targets: {len(all_targets)}")

    if not all_targets:
        print("No formula cells found. Exiting.")
        return

    print("\n2. Building dependency graph...", flush=True)
    t_build = time.perf_counter()
    print("   Starting create_dependency_graph...", flush=True)
    try:
        graph = create_graph_for_targets(all_targets)
    except DynamicRefError as e:
        print(f"\n   DynamicRefError: {e}")
        print(
            "   This profile script uses empty constraints and cached dynamic refs only."
            " For constraint-based resolution use scripts/map_lic_dsf.py."
        )
        raise

    build_s = time.perf_counter() - t_build
    print(f"   Graph build time: {build_s:.2f}s")

    print(f"   Nodes in graph: {len(graph)}")
    print(f"   Leaf nodes: {sum(1 for _ in graph.leaves())}")
    print(f"   Formula nodes: {len(graph) - sum(1 for _ in graph.leaves())}")

    sheets: dict[str, int] = {}
    for key in graph:
        node = graph.get_node(key)
        if node:
            sheets[node.sheet] = sheets.get(node.sheet, 0) + 1

    print("\n   Nodes by sheet:")
    for sheet_name in sorted(sheets.keys()):
        print(f"      {sheet_name}: {sheets[sheet_name]}")

    print("\n3. Workbook calculation settings...")
    settings = get_calc_settings(WORKBOOK_PATH)
    print(f"   Iterate enabled: {settings.iterate_enabled}")
    print(f"   Iterate count:   {settings.iterate_count}")
    print(f"   Iterate delta:   {settings.iterate_delta}")

    print("\n4. Cycle analysis...")
    report = graph.cycle_report()
    print(f"   Must-cycles: {len(report.must_cycles)}")
    print(f"   May-cycles:  {len(report.may_cycles)}")
    if report.example_must_cycle_path:
        print(
            f"   Example must-cycle path: {' -> '.join(report.example_must_cycle_path)}"
        )
    if report.example_may_cycle_path:
        print(
            f"   Example may-cycle path:  {' -> '.join(report.example_may_cycle_path)}"
        )

    print("\n5. Validating against calcChain.xml...")
    scope = {parse_range_spec(entry["range_spec"])[0] for entry in EXPORT_RANGES}
    result = validate_graph(graph, WORKBOOK_PATH, scope=scope)

    print(f"   Valid: {result.is_valid}")
    for msg in result.messages:
        print(f"   {msg}")

    if result.in_graph_not_in_chain:
        print(
            f"\n   Cells in graph but not in calcChain ({len(result.in_graph_not_in_chain)}):"
        )
        for cell in sorted(result.in_graph_not_in_chain)[:10]:
            print(f"      {cell}")
        if len(result.in_graph_not_in_chain) > 10:
            print(f"      ... and {len(result.in_graph_not_in_chain) - 10} more")

    print("\n6. Computing evaluation order...")
    try:
        order = graph.evaluation_order(strict=False)
        print(f"   Evaluation order computed: {len(order)} nodes")
        print(f"   First 5 (leaves): {order[:5]}")
        print(f"   Last 5 (targets): {order[-5:]}")
    except CycleError as e:
        kind = "must-cycle" if e.is_must_cycle else "may-cycle"
        print(f"   Error ({kind}): {e}")
        if e.cycle_path:
            print(f"   Cycle path: {' -> '.join(e.cycle_path)}")

    print("\n7. Sample visualization (first target's immediate deps)...")
    if all_targets:
        sample_target = all_targets[0]
        sample_deps = graph.dependencies(sample_target)
        print(f"   {sample_target} depends on {len(sample_deps)} cells:")
        for dep in sorted(sample_deps)[:5]:
            guard = graph.edge_attrs(sample_target, dep).get("guard")
            if guard is None:
                print(f"      {dep}")
            else:
                print(f"      {dep}  [guarded: {guard}]")
        if len(sample_deps) > 5:
            print(f"      ... and {len(sample_deps) - 5} more")

        try:
            dot = to_graphviz(graph, highlight={sample_target}, rankdir="LR")
            print("\n   GraphViz DOT (truncated to first ~40 lines):")
            for line in dot.splitlines()[:40]:
                print(f"      {line}")
            if len(dot.splitlines()) > 40:
                print("      ...")
        except Exception as e:
            print(f"   Could not render GraphViz DOT: {e}")

    print("\n" + "=" * 70)
    print("Done.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="LIC-DSF dependency mapping (profiling build).",
    )
    parser.add_argument(
        "--profile-post-load",
        type=Path,
        nargs="?",
        const=Path("map_post_load.prof"),
        default=None,
        metavar="OUT.prof",
        help=(
            "Pre-load the workbook, patch fastpyxl.load_workbook, then run cProfile "
            "only around create_dependency_graph (excludes parse/load from stats). "
            "Writes OUT.prof (default: map_post_load.prof)."
        ),
    )
    parser.add_argument(
        "--profile-dump-on-sigterm",
        action="store_true",
        help="On SIGTERM, dump stats then exit 124 (use with timeout(1) so the .prof is written).",
    )
    args = parser.parse_args()
    if args.profile_post_load is not None:
        profile_post_load_graph(
            output_prof=args.profile_post_load,
            dump_on_sigterm=args.profile_dump_on_sigterm,
        )
    else:
        main()
