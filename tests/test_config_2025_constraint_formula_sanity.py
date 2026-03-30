from __future__ import annotations

from pathlib import Path

import pytest

from src.configs import load_template_config


@pytest.mark.slow
def test_2025_checked_sheet_constraints_are_not_formula_cells() -> None:
    """Every constrained cell must be a leaf (non-formula) in the template workbook."""
    cfg = load_template_config("2025-08-12")
    wb = Path.cwd() / cfg.WORKBOOK_PATH
    if not wb.is_file():
        pytest.skip("Template workbook not present (e.g. not checked into the repo)")
    cfg.verify_lic_dsf_constraints_target_leaves(wb, cfg.LicDsfConstraints)
