from __future__ import annotations

from typing import Annotated, get_args, get_origin, get_type_hints

from excel_grapher.core.cell_types import RealBetween
from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_market_financing_e53_g53_have_financial_domain_for_dynamic_refs() -> None:
    """C4 E:G row 51–53 are formulas on-tab; INDIRECT paths use codename `Market_financing!`."""
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    for coord in ("E53", "G53"):
        alias = format_key("Market_financing", coord)
        ann = hints[alias]
        assert get_origin(ann) is Annotated
        assert any(isinstance(m, RealBetween) for m in get_args(ann)[1:])
