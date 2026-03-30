#!/usr/bin/env python3
"""
First-pass diagnostic for LIC-DSF full dependency graph builds.

- Prints periodic main-thread stack traces to stderr (faulthandler) so a slow or
  stuck run still produces output without py-spy.
- Optional --no-provenance A/B: same graph as lic_dsf_pipeline.build_graph except
  capture_dependency_provenance=False (isolates provenance merge cost).

Run from repo root with a wall timeout, e.g.:

  uv run python scripts/diagnose_build_graph.py 2> diagnose_build_graph.stderr
  timeout 120 uv run python scripts/diagnose_build_graph.py
  timeout 120 uv run python scripts/diagnose_build_graph.py --no-provenance
  timeout 60 uv run python scripts/diagnose_build_graph.py --one-target "PV_Base!A1"

First-pass interpretation of periodic stacks: the process advances through
``fastpyxl.load_workbook`` (xlsm parse), then ``create_dependency_graph`` with
most samples in ``extract_expr_deps`` → ``normalize_formula`` / ``format_cell_key``
(``excel_grapher.grapher.parser``). With ``capture_dependency_provenance=True``,
stacks sometimes include ``provenance_collect.collect_provenance_for_formula``.
That matches a very large graph build, not a single stuck instruction.
"""

from __future__ import annotations

import argparse
import faulthandler
import sys
import time
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))


def _main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--config",
        default="2025-08-12",
        help="Template id passed to load_template_config (default: 2025-08-12)",
    )
    parser.add_argument(
        "--interval",
        type=float,
        default=15.0,
        metavar="SEC",
        help="Print main-thread stack to stderr every SEC seconds (default: 15)",
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=50,
        help="Graph max_depth (default: 50)",
    )
    parser.add_argument(
        "--no-provenance",
        action="store_true",
        help="Use capture_dependency_provenance=False (else match build_graph)",
    )
    parser.add_argument(
        "--one-target",
        metavar="ADDR",
        default="",
        help="If set, graph only this sheet-qualified address (minimal repro helper)",
    )
    args = parser.parse_args()

    faulthandler.enable(all_threads=False)
    if args.interval > 0:
        faulthandler.dump_traceback_later(args.interval, repeat=True, file=sys.stderr)

    # Import after argparse so --help works without template on PYTHONPATH cwd issues
    from excel_grapher.grapher import create_dependency_graph

    from src.configs import load_template_config
    from src.lic_dsf_config import discover_targets_from_ranges

    cfg = load_template_config(args.config)
    wb = Path(cfg.WORKBOOK_PATH)
    if not wb.is_file():
        wb = Path.cwd() / wb
    if not wb.is_file():
        print(f"error: workbook missing: {wb}", file=sys.stderr)
        sys.exit(2)

    if args.one_target:
        targets = [args.one_target.strip()]
    else:
        targets = discover_targets_from_ranges(cfg.EXPORT_RANGES)
    dr = cfg.get_dynamic_ref_config()
    t0 = time.perf_counter()
    print(
        f"[diagnose_build_graph] workbook={wb} targets={len(targets)} "
        f"max_depth={args.max_depth} provenance={not args.no_provenance}",
        file=sys.stderr,
        flush=True,
    )

    graph = create_dependency_graph(
        wb,
        targets,
        load_values=False,
        max_depth=args.max_depth,
        dynamic_refs=dr,
        use_cached_dynamic_refs=False,
        capture_dependency_provenance=not args.no_provenance,
    )
    graph.compress_identity_transits()
    elapsed = time.perf_counter() - t0
    print(
        f"[diagnose_build_graph] OK nodes={len(graph)} elapsed_s={elapsed:.3f}",
        file=sys.stderr,
        flush=True,
    )


if __name__ == "__main__":
    _main()
