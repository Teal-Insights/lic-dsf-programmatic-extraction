from __future__ import annotations

from typing import Annotated, get_args, get_origin, get_type_hints

from excel_grapher import NotEqualCell
from excel_grapher.core.cell_types import Between
from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_pv_lc_nr_maturity_mirrors_cannot_equal_grace_mirrors() -> None:
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)

    for sheet in ("PV_LC_NR1", "PV_LC_NR2", "PV_LC_NR3"):
        for maturity_row in range(31, 412, 19):
            ann = hints[format_key(sheet, f"B{maturity_row}")]
            assert get_origin(ann) is Annotated
            base, *metadata = get_args(ann)
            assert base == (int | None)
            assert Between(1, 100) in metadata
            assert NotEqualCell(format_key(sheet, f"B{maturity_row - 1}")) in metadata
