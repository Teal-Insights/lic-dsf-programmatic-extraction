"""Tests for scripts/benchmark_graph_provenance.py root-scope timing wrapper."""

from __future__ import annotations

import importlib.util
import sys
import time
from pathlib import Path
from typing import Any

_ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location(
    "benchmark_graph_provenance",
    _ROOT / "scripts" / "benchmark_graph_provenance.py",
)
assert _spec and _spec.loader
_mod = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _mod
_spec.loader.exec_module(_mod)
_wrap = _mod.wrap_collect_provenance_for_root_timing


def test_root_timing_one_row_for_recursive_orig() -> None:
    """Recursive orig through wrapped: single root row, elapsed covers full tree."""
    w_holder: list[Any] = []

    def orig(
        formula: str,
        *,
        normalized_formula: str | None,
        current_sheet: str,
        current_a1: str,
    ) -> dict[str, str]:
        if formula == "=X":
            time.sleep(0.01)
            return {}
        time.sleep(0.01)
        return w_holder[0](
            "=X",
            normalized_formula=None,
            current_sheet=current_sheet,
            current_a1=current_a1,
        )

    w, roots = _wrap(orig, lambda s, a: f"{s}!{a}")
    w_holder.append(w)
    w("=ROOT", normalized_formula=None, current_sheet="Sh", current_a1="B2")

    assert len(roots) == 1
    sec, addr, nch = roots[0]
    assert addr == "Sh!B2"
    assert nch == len("=ROOT")
    assert sec >= 0.02


def test_root_timing_separate_rows_for_sequential_calls() -> None:
    def orig(
        formula: str,
        *,
        normalized_formula: str | None,
        current_sheet: str,
        current_a1: str,
    ) -> dict[str, str]:
        return {}

    w, roots = _wrap(orig, lambda s, a: f"{a}")
    w("=A", normalized_formula=None, current_sheet="S", current_a1="A1")
    w("=BB", normalized_formula=None, current_sheet="S", current_a1="A2")

    assert len(roots) == 2
    assert roots[0][2] == 2
    assert roots[1][2] == 3
