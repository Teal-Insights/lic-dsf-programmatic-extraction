from __future__ import annotations

from dataclasses import dataclass

import openpyxl.utils.cell

from lic_dsf_input_setters import (
    WideYearSeriesSpec,
    apply_year_series_array,
    apply_year_series_mapping,
    build_wide_year_series_spec,
)


@dataclass
class DummyCtx:
    inputs: dict[str, object]

    def set_inputs(self, inputs: dict[str, object]) -> None:
        self.inputs.update(inputs)


def test_build_wide_year_series_spec_ext_debt_data_row10() -> None:
    # Known year header row around Ext_Debt_Data!E10:H10
    spec = build_wide_year_series_spec(
        workbook_path="workbooks/lic-dsf-template.xlsm",
        sheet="Ext_Debt_Data",
        row=10,
        start_col=openpyxl.utils.cell.column_index_from_string("E"),
        end_col=openpyxl.utils.cell.column_index_from_string("H"),
    )

    assert spec.years == (2023, 2024, 2025, 2026)
    assert spec.year_to_address[2023] == "Ext_Debt_Data!E10"
    assert spec.year_to_address[2026] == "Ext_Debt_Data!H10"


def test_apply_year_series_mapping_sets_only_provided_years_and_none_to_zero() -> None:
    spec = WideYearSeriesSpec(
        years=(2023, 2024, 2025),
        year_to_address={2023: "S!A1", 2024: "S!B1", 2025: "S!C1"},
    )
    ctx = DummyCtx(inputs={})

    assignment = apply_year_series_mapping(
        ctx,
        spec,
        {2023: 1.0, 2025: None},
        strict=True,
    )

    assert ctx.inputs == {"S!A1": 1.0, "S!C1": 0}
    assert assignment.applied == {2023: "S!A1", 2025: "S!C1"}
    assert assignment.ignored == {}
    assert assignment.years == (2023, 2024, 2025)


def test_apply_year_series_mapping_strict_false_ignores_unknown_years() -> None:
    spec = WideYearSeriesSpec(
        years=(2023, 2024),
        year_to_address={2023: "S!A1", 2024: "S!B1"},
    )
    ctx = DummyCtx(inputs={})

    assignment = apply_year_series_mapping(
        ctx,
        spec,
        {2022: 9, 2023: 1},
        strict=False,
    )

    assert ctx.inputs == {"S!A1": 1}
    assert assignment.applied == {2023: "S!A1"}
    assert assignment.ignored == {2022: 9}


def test_apply_year_series_array_requires_contiguous_years() -> None:
    spec = WideYearSeriesSpec(
        years=(2023, 2025),
        year_to_address={2023: "S!A1", 2025: "S!B1"},
    )
    ctx = DummyCtx(inputs={})

    try:
        apply_year_series_array(ctx, spec, [1, 2], start_year=2023)
    except ValueError as e:
        assert "Non-contiguous years" in str(e)
    else:
        raise AssertionError("Expected ValueError for non-contiguous years")

