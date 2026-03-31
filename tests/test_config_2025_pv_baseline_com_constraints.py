from __future__ import annotations

from typing import Annotated, get_args, get_origin, get_type_hints

from excel_grapher.core.cell_types import RealBetween
from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_pv_baseline_com_bc49_in_ar_bp_band() -> None:
    """Debt-service year grid extends through BP; template leaves are non-formulas in AR:BP."""
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    key = format_key("PV_baseline_com", "BC49")
    ann = hints[key]
    assert get_origin(ann) is Annotated
    assert any(isinstance(m, RealBetween) for m in get_args(ann)[1:])
