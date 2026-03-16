from __future__ import annotations

from pathlib import Path

from src.configs import load_template_config
from src.lic_dsf_export import generate_setter_method_name, generate_setters_module, load_input_groups

_cfg = load_template_config("2026-01-31")
_INPUT_GROUPS_PATH = Path(__file__).resolve().parents[1] / "src" / "configs" / "2026-01-31" / "input_groups.json"
_WORKBOOK_PATH = _cfg.WORKBOOK_PATH


def _pick_wide_year_series(groups: list[dict]) -> dict:
    for g in groups:
        shape = g.get("shape") or {}
        if (
            g.get("mode") == "ignore_column_axis_years"
            and shape.get("rows") == 1
            and shape.get("cols", 0) > 1
            and g.get("cells")
            and g.get("range_a1")
        ):
            return g
    raise AssertionError("No wide year-series group found in input_groups.json")


def _pick_non_year_group(groups: list[dict]) -> dict:
    for g in groups:
        shape = g.get("shape") or {}
        if (
            g.get("mode") == "no_year_axis"
            and shape.get("rows")
            and shape.get("cols")
            and g.get("cells")
            and g.get("range_a1")
        ):
            return g
    raise AssertionError("No non-year group found in input_groups.json")


def test_generate_setters_module_contains_wide_year_series_setter() -> None:
    groups = load_input_groups(_INPUT_GROUPS_PATH)
    target = _pick_wide_year_series(groups)
    module = generate_setters_module(workbook=_WORKBOOK_PATH, groups=[target])

    assert "class LicDsfContext(EvalContext):" in module
    name = generate_setter_method_name(
        str(target.get("sheet")),
        list(target.get("row_labels_key") or []),
        str(target.get("group_id", "group")),
    )
    assert f"def {name}" in module
    assert f"def {name}_from_array" not in module
    cells = sorted(str(c) for c in target.get("cells", []))
    assert cells[0] in module
    assert cells[-1] in module


def test_generate_setters_module_contains_no_year_range_setters() -> None:
    groups = load_input_groups(_INPUT_GROUPS_PATH)
    target = _pick_non_year_group(groups)
    module = generate_setters_module(workbook=_WORKBOOK_PATH, groups=[target])

    assert "class RangeAssignment" in module
    name = generate_setter_method_name(
        str(target.get("sheet")),
        list(target.get("row_labels_key") or []),
        str(target.get("group_id", "group")),
    )
    assert f"def {name}" in module
    cells = sorted(str(c) for c in target.get("cells", []))
    assert cells[0] in module


def test_generate_setters_module_includes_workbook_loader() -> None:
    groups = load_input_groups(_INPUT_GROUPS_PATH)
    target = _pick_wide_year_series(groups)
    module = generate_setters_module(workbook=_WORKBOOK_PATH, groups=[target])

    assert "def load_inputs_from_workbook" in module
    assert "_read_inputs_from_workbook" in module
    assert "DEFAULT_INPUTS" in module
