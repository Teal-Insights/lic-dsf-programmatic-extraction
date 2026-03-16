from __future__ import annotations

from pathlib import Path

import pytest

from src.configs import load_template_config
from src.lic_dsf_export import generate_setter_method_name, generate_setters_module, load_input_groups

_cfg = load_template_config("2026-01-31")
_INPUT_GROUPS_PATH = Path(__file__).resolve().parents[1] / "src" / "configs" / "2026-01-31" / "input_groups.json"
_WORKBOOK_PATH = _cfg.WORKBOOK_PATH


def _pick_tall_year_group(groups: list[dict]) -> dict:
    for g in groups:
        if g.get("mode") == "ignore_row_axis_years" and g.get("cells") and g.get("range_a1"):
            return g
    raise pytest.skip("No tall year-series groups found in input_groups.json")


def test_codegen_generates_tall_year_row_setter() -> None:
    groups = load_input_groups(_INPUT_GROUPS_PATH)
    g = _pick_tall_year_group(groups)

    module = generate_setters_module(workbook=_WORKBOOK_PATH, groups=[g])
    assert "class YearRowAssignment" in module
    base = generate_setter_method_name(
        str(g.get("sheet")),
        list(g.get("row_labels_key") or []),
        str(g.get("group_id", "group")),
    )
    assert f"def {base}_by_year" in module
    assert f"def {base}_by_year_from_array" not in module
