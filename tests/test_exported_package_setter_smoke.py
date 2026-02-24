from __future__ import annotations

import sys
from pathlib import Path

import pytest

from lic_dsf_export import generate_setter_method_name, load_input_groups
from lic_dsf_input_setters import BASE_YEAR_ADDRESS, build_wide_year_series_spec


def _pick_wide_year_series(groups: list[dict]) -> dict:
    for g in groups:
        shape = g.get("shape") or {}
        if (
            g.get("mode") == "ignore_column_axis_years"
            and shape.get("rows") == 1
            and shape.get("cols", 0) > 1
            and g.get("bounding_box")
        ):
            return g
    raise AssertionError("No wide year-series group found in input_groups.json")


def _normalize_address(address: str) -> str:
    if address.startswith("'") and "'!" in address:
        sheet, rest = address.split("!", 1)
        return f"{sheet[1:-1]}!{rest}"
    return address


@pytest.mark.xfail(
    reason="Requires regenerated export package with offset-based setters (INDICATOR_CONFIG update needed)"
)
def test_exported_package_has_year_series_setter() -> None:
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    groups = load_input_groups(Path("input_groups.json"))
    target = _pick_wide_year_series(groups)
    bbox = target.get("bounding_box") or {}
    sheet = str(target.get("sheet"))
    row = int(bbox.get("start_row"))
    start_col = int(bbox.get("start_col"))
    end_col = int(bbox.get("end_col"))

    spec = build_wide_year_series_spec(
        workbook_path="workbooks/lic-dsf-template-2026-01-31.xlsm",
        sheet=sheet,
        row=row,
        start_col=start_col,
        end_col=end_col,
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
        pytest.skip("Non-contiguous offsets; array mapping requires contiguous offsets.")

    ctx2 = lic_dsf.make_context()
    assignment2 = getattr(ctx2, name)([10, 20], start_year=offsets[0], strict=True)
    assert assignment2.applied[offsets[0]] in ctx2.inputs


def test_exported_package_exports_range_assignment() -> None:
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "RangeAssignment")
    assert "RangeAssignment" in getattr(lic_dsf, "__all__", [])


def test_exported_package_exports_context_with_setters() -> None:
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "LicDsfContext")
    assert "LicDsfContext" in getattr(lic_dsf, "__all__", [])


def test_exported_package_exports_year_row_assignment() -> None:
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "YearRowAssignment")
    assert "YearRowAssignment" in getattr(lic_dsf, "__all__", [])
