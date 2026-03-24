from __future__ import annotations

import sys
from pathlib import Path

import pytest

from src.configs import load_template_config
from src.lic_dsf_export import generate_setter_method_name, load_input_groups
from src.lic_dsf_input_setters import build_wide_year_series_spec

_cfg = load_template_config("2026-01-31")
_INPUT_GROUPS_PATH = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "configs"
    / "2026-01-31"
    / "input_groups.json"
)
_WORKBOOK_PATH = _cfg.WORKBOOK_PATH
_EXPORT_DIR = _cfg.EXPORT_DIR
_REGION = _cfg.REGION_CONFIG

_EXPORT_PKG_INIT = _EXPORT_DIR / "lic_dsf_2026_01_31" / "__init__.py"

pytestmark = pytest.mark.skipif(
    not _EXPORT_PKG_INIT.is_file(),
    reason="Run the export pipeline to populate dist/ before this integration test.",
)


def _pick_wide_year_series(groups: list[dict]) -> dict:
    for g in groups:
        shape = g.get("shape") or {}
        if (
            g.get("mode") == "ignore_column_axis_years"
            and shape.get("rows") == 1
            and shape.get("cols", 0) >= 10
            and g.get("bounding_box")
        ):
            return g
    raise AssertionError("No wide year-series group found in input_groups.json")


def _normalize_address(address: str) -> str:
    if address.startswith("'") and "'!" in address:
        sheet, rest = address.split("!", 1)
        return f"{sheet[1:-1]}!{rest}"
    return address


def test_exported_package_has_year_series_setter() -> None:
    sys.path.insert(0, str(_EXPORT_DIR.resolve()))
    import lic_dsf_2026_01_31 as lic_dsf  # type: ignore

    groups = load_input_groups(_INPUT_GROUPS_PATH)
    target = _pick_wide_year_series(groups)
    bbox = target.get("bounding_box") or {}
    sheet = str(target.get("sheet"))
    row = bbox.get("start_row")
    start_col = bbox.get("start_col")
    end_col = bbox.get("end_col")
    if not isinstance(row, int):
        pytest.skip("Missing bounding box metadata for wide year series")
    if not isinstance(start_col, int):
        pytest.skip("Missing bounding box metadata for wide year series")
    if not isinstance(end_col, int):
        pytest.skip("Missing bounding box metadata for wide year series")
    row = row
    start_col = start_col
    end_col = end_col

    spec = build_wide_year_series_spec(
        workbook_path=str(_WORKBOOK_PATH),
        sheet=sheet,
        row=row,
        start_col=start_col,
        end_col=end_col,
        region_config=_REGION,
    )
    name = generate_setter_method_name(
        sheet,
        list(target.get("row_labels_key") or []),
        str(target.get("group_id", "group")),
    )

    ctx = lic_dsf.make_context()

    # Use raw offset keys (no base year needed)
    offset = spec.offsets[0]
    assignment = getattr(ctx, name)({offset: 123, spec.offsets[-1]: None}, strict=True)

    assert _normalize_address(assignment.applied[offset]) == _normalize_address(
        spec.offset_to_address[offset]
    )
    assert ctx.inputs[assignment.applied[offset]] == 123
    assert ctx.inputs[assignment.applied[spec.offsets[-1]]] == 0

    offsets = list(spec.offsets)
    expected = list(range(offsets[0], offsets[0] + len(offsets)))
    if offsets != expected:
        pytest.skip(
            "Non-contiguous offsets; array mapping requires contiguous offsets."
        )

    ctx2 = lic_dsf.make_context()
    assignment2 = getattr(ctx2, name)([10, 20], start_year=offsets[0], strict=True)
    assert assignment2.applied[offsets[0]] in ctx2.inputs


def test_exported_package_exports_range_assignment() -> None:
    sys.path.insert(0, str(_EXPORT_DIR.resolve()))
    import lic_dsf_2026_01_31 as lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "RangeAssignment")
    assert "RangeAssignment" in getattr(lic_dsf, "__all__", [])


def test_exported_package_exports_context_with_setters() -> None:
    sys.path.insert(0, str(_EXPORT_DIR.resolve()))
    import lic_dsf_2026_01_31 as lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "LicDsfContext")
    assert "LicDsfContext" in getattr(lic_dsf, "__all__", [])


def test_exported_package_exports_year_row_assignment() -> None:
    sys.path.insert(0, str(_EXPORT_DIR.resolve()))
    import lic_dsf_2026_01_31 as lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "YearRowAssignment")
    assert "YearRowAssignment" in getattr(lic_dsf, "__all__", [])
