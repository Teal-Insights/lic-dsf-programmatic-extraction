from __future__ import annotations

from pathlib import Path

from lic_dsf_export import generate_setters_module, load_input_groups


def test_codegen_generates_tall_year_row_setter_for_local_debt_financing() -> None:
    groups = load_input_groups(Path("input_groups_export.json"))
    # Pick a tall-format group with a stable non-year label.
    g = next(
        x
        for x in groups
        if x.get("mode") == "ignore_row_axis_years"
        and x.get("sheet") == "Input 5 - Local-debt Financing"
        and (x.get("row_labels_key") or [])[:1] == ["Central bank financing"]
    )

    module = generate_setters_module(workbook=Path("workbooks/lic-dsf-template.xlsm"), groups=[g])
    assert "class YearRowAssignment" in module
    assert "def set_input_5_local_debt_financing_central_bank_financing_by_year" in module

