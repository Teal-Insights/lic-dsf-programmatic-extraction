from __future__ import annotations

import sys
from pathlib import Path


def test_exported_package_has_year_series_setter() -> None:
    # Import the generated package from ./export without requiring installation.
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    ctx = lic_dsf.make_context()
    # Known setter generated from input_groups_export.json group Ext_Debt_Data!E10:H10
    assignment = ctx.set_ext_debt_data_external_debt_excluding_locally_issued_debt(
        {2023: 123, 2026: None},
        strict=True,
    )

    assert assignment.applied[2023] == "Ext_Debt_Data!E10"
    assert ctx.inputs["Ext_Debt_Data!E10"] == 123
    assert ctx.inputs["Ext_Debt_Data!H10"] == 0


def test_exported_package_exports_range_assignment() -> None:
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "RangeAssignment")
    assert "RangeAssignment" in getattr(lic_dsf, "__all__", [])


def test_exported_package_exports_year_row_assignment() -> None:
    sys.path.insert(0, str(Path("export").resolve()))
    import lic_dsf  # type: ignore

    assert hasattr(lic_dsf, "YearRowAssignment")
    assert "YearRowAssignment" in getattr(lic_dsf, "__all__", [])

