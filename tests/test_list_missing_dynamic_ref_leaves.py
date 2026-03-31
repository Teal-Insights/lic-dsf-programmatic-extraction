from __future__ import annotations

from pathlib import Path

import pytest


def test_list_missing_dynamic_ref_leaves_forwards_workbook_and_kwargs(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict = {}

    def fake(
        workbook: object,
        targets: object,
        *,
        dynamic_refs: object = None,
        max_depth: int = 50,
        max_range_cells: int = 5000,
    ) -> list[str]:
        captured["workbook"] = workbook
        captured["targets"] = list(targets)
        captured["dynamic_refs"] = dynamic_refs
        captured["max_depth"] = max_depth
        captured["max_range_cells"] = max_range_cells
        return ["'S'!A1"]

    monkeypatch.setattr(
        "src.lic_dsf_pipeline.list_dynamic_ref_constraint_candidates",
        fake,
    )
    from src.lic_dsf_pipeline import list_missing_dynamic_ref_leaves

    p = Path("/tmp/fake.xlsm")
    out = list_missing_dynamic_ref_leaves(
        p,
        ["S!B2"],
        12,
        wb_formulas=None,
        dynamic_refs=None,
        max_range_cells=99,
    )
    assert out == ["'S'!A1"]
    assert captured["workbook"] is p
    assert captured["targets"] == ["S!B2"]
    assert captured["dynamic_refs"] is None
    assert captured["max_depth"] == 12
    assert captured["max_range_cells"] == 99
