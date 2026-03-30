from __future__ import annotations

from pathlib import Path

import pytest

from src.configs import load_template_config


@pytest.mark.slow
def test_2025_checked_sheet_constraints_are_not_formula_cells() -> None:
    """Sheets outside the PV/COM/input-calculation set must not constrain formula cells.

    Exception: START!L10 is VLOOKUP-backed; the graph treats its domain via constraints
    because the evaluator does not implement VLOOKUP (see config comments).
    """
    cfg = load_template_config("2025-08-12")
    wb = Path.cwd() / cfg.WORKBOOK_PATH
    if not wb.is_file():
        pytest.skip("Template workbook not present (e.g. not checked into the repo)")
    cfg._assert_checked_sheet_constraint_cells_are_not_formulas(cfg.LicDsfConstraints, wb)
