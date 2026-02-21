from __future__ import annotations

from pathlib import Path

import pytest

from lic_dsf_export import generate_setter_method_name, generate_setters_module, load_input_groups


def _pick_tall_year_group(groups: list[dict]) -> dict:
    for g in groups:
        if g.get("mode") == "ignore_row_axis_years" and g.get("cells") and g.get("range_a1"):
            return g
    raise pytest.skip("No tall year-series groups found in input_groups.json")


def test_codegen_generates_tall_year_row_setter() -> None:
    groups = load_input_groups(Path("input_groups.json"))
    g = _pick_tall_year_group(groups)

    module = generate_setters_module(workbook=Path("workbooks/lic-dsf-template-2026-01-31.xlsm"), groups=[g])
    assert "class YearRowAssignment" in module
    base = generate_setter_method_name(
        str(g.get("sheet")),
        list(g.get("row_labels_key") or []),
        str(g.get("group_id", "group")),
    )
    assert f"def {base}_by_year" in module
    assert f"def {base}_by_year_from_array" not in module

