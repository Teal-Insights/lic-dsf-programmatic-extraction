from __future__ import annotations

from typing import Annotated, get_args, get_origin, get_type_hints

from excel_grapher.core.cell_types import GreaterThanCell, constraints_to_cell_type_env
from excel_grapher.core.formula_ast import BinaryOpNode, parse as parse_ast
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    _FiniteInts,
    _IntBounds,
    _refine_difference_domain,
    format_key,
)

from src.configs import load_template_config


def test_2025_input4_h10_includes_greater_than_g10_metadata() -> None:
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    key = format_key("Input 4 - External Financing", "H10")
    ann = hints[key]
    assert get_origin(ann) is Annotated
    metas = list(get_args(ann)[1:])
    assert any(isinstance(m, GreaterThanCell) for m in metas)
    gtc = next(m for m in metas if isinstance(m, GreaterThanCell))
    assert gtc.other == format_key("Input 4 - External Financing", "G10")


def test_2025_input4_h10_minus_g10_domain_refined() -> None:
    cfg = load_template_config("2025-08-12")
    env = constraints_to_cell_type_env(cfg.LicDsfConstraints, {})
    ast = parse_ast("='Input 4 - External Financing'!H10-'Input 4 - External Financing'!G10")
    assert isinstance(ast, BinaryOpNode)
    limits = DynamicRefLimits()
    wide = _IntBounds(-100, 100)
    refined = _refine_difference_domain(ast, env, wide, limits)
    assert refined is not None
    if isinstance(refined, _IntBounds):
        assert refined.lo >= 1
    else:
        assert isinstance(refined, _FiniteInts)
        assert min(refined.values) >= 1
