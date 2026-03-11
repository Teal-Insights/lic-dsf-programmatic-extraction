from __future__ import annotations

from dataclasses import dataclass, field

import openpyxl.utils.cell

from src.lic_dsf_input_setters import (
    BASE_YEAR_ADDRESS,
    WideYearSeriesSpec,
    apply_year_series_array,
    apply_year_series_mapping,
    build_wide_year_series_spec,
)


@dataclass
class DummyCtx:
    inputs: dict[str, object] = field(default_factory=dict)

    def set_inputs(self, inputs: dict[str, object]) -> None:
        self.inputs.update(inputs)


def test_build_wide_year_series_spec_ext_debt_data_row10() -> None:
    spec = build_wide_year_series_spec(
        workbook_path="workbooks/lic-dsf-template-2026-01-31.xlsm",
        sheet="Ext_Debt_Data",
        row=10,
        start_col=openpyxl.utils.cell.column_index_from_string("E"),
        end_col=openpyxl.utils.cell.column_index_from_string("H"),
    )

    assert spec.offsets == (-1, 0, 1, 2)
    assert spec.offset_to_address[-1] == "Ext_Debt_Data!E10"
    assert spec.offset_to_address[2] == "Ext_Debt_Data!H10"


def test_apply_year_series_mapping_with_offset_keys() -> None:
    spec = WideYearSeriesSpec(
        offsets=(-1, 0, 1),
        offset_to_address={-1: "S!A1", 0: "S!B1", 1: "S!C1"},
    )
    ctx = DummyCtx()

    assignment = apply_year_series_mapping(
        ctx,
        spec,
        {-1: 1.0, 1: None},
        strict=True,
    )

    assert ctx.inputs == {"S!A1": 1.0, "S!C1": 0}
    assert assignment.applied == {-1: "S!A1", 1: "S!C1"}
    assert assignment.ignored == {}
    assert assignment.offsets == (-1, 0, 1)


def test_apply_year_series_mapping_with_year_keys() -> None:
    """Year-like keys (1900-2100) are resolved via the base year in ctx.inputs."""
    spec = WideYearSeriesSpec(
        offsets=(0, 1, 2),
        offset_to_address={0: "S!A1", 1: "S!B1", 2: "S!C1"},
    )
    ctx = DummyCtx(inputs={BASE_YEAR_ADDRESS: 2024})

    assignment = apply_year_series_mapping(
        ctx,
        spec,
        {2024: 10, 2026: 20},
        strict=True,
    )

    assert ctx.inputs["S!A1"] == 10
    assert ctx.inputs["S!C1"] == 20
    assert assignment.applied == {2024: "S!A1", 2026: "S!C1"}


def test_apply_year_series_mapping_strict_false_ignores_unknown() -> None:
    spec = WideYearSeriesSpec(
        offsets=(0, 1),
        offset_to_address={0: "S!A1", 1: "S!B1"},
    )
    ctx = DummyCtx()

    assignment = apply_year_series_mapping(
        ctx,
        spec,
        {-1: 9, 0: 1},
        strict=False,
    )

    assert ctx.inputs == {"S!A1": 1}
    assert assignment.applied == {0: "S!A1"}
    assert assignment.ignored == {-1: 9}


def test_apply_year_series_array_requires_contiguous_offsets() -> None:
    spec = WideYearSeriesSpec(
        offsets=(0, 2),
        offset_to_address={0: "S!A1", 2: "S!B1"},
    )
    ctx = DummyCtx()

    try:
        apply_year_series_array(ctx, spec, [1, 2], start_year=0)
    except ValueError as e:
        assert "Non-contiguous" in str(e)
    else:
        raise AssertionError("Expected ValueError for non-contiguous offsets")


def test_apply_year_series_mapping_year_key_without_base_year_raises() -> None:
    spec = WideYearSeriesSpec(
        offsets=(0, 1),
        offset_to_address={0: "S!A1", 1: "S!B1"},
    )
    ctx = DummyCtx()

    try:
        apply_year_series_mapping(ctx, spec, {2024: 10}, strict=True)
    except ValueError as e:
        assert "base year" in str(e).lower()
    else:
        raise AssertionError("Expected ValueError when using year key without base year")
