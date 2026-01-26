from __future__ import annotations

import json
from pathlib import Path

from lic_dsf_export import generate_setters_module, load_input_groups


def test_generate_setters_module_contains_ext_debt_data_row10_setter() -> None:
    groups = load_input_groups(Path("input_groups_export.json"))
    # Reduce to a single known year-series group to keep the generated module stable for assertions.
    target = next(g for g in groups if g.get("range_a1") == "Ext_Debt_Data!E10:H10")
    module = generate_setters_module(workbook=Path("workbooks/lic-dsf-template.xlsm"), groups=[target])

    assert "class LicDsfContext(EvalContext):" in module
    # Naming scheme is deterministic: set_<sheet>_<label>
    assert "def set_ext_debt_data_external_debt_excluding_locally_issued_debt" in module
    # Year mapping should be embedded.
    assert "Ext_Debt_Data!E10" in module
    assert "Ext_Debt_Data!H10" in module


def test_generate_setters_module_contains_no_year_range_setters() -> None:
    groups = load_input_groups(Path("input_groups_export.json"))
    scalar = next(g for g in groups if g.get("range_a1") == "B1_GDP_ext!AF4:AF4")
    row_vec = next(g for g in groups if g.get("range_a1") == "Ext_Debt_Data!BP79:CC79")
    table = next(g for g in groups if g.get("range_a1") == "PV_Base!B40:C44")

    module = generate_setters_module(
        workbook=Path("workbooks/lic-dsf-template.xlsm"),
        groups=[scalar, row_vec, table],
    )

    assert "class RangeAssignment" in module
    assert "def set_ext_debt_data_ida_new_60_year_credits" in module
    assert "Ext_Debt_Data!BP79" in module
    assert "PV_Base!B40" in module

