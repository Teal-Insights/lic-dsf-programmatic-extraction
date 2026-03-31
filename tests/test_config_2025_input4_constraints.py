from __future__ import annotations

from typing import Literal, get_args, get_origin, get_type_hints

from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_input4_g29_h29_literal_pair_nonzero_span() -> None:
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    g29 = hints[format_key("Input 4 - External Financing", "G29")]
    h29 = hints[format_key("Input 4 - External Financing", "H29")]
    assert get_origin(g29) is Literal and get_args(g29) == (1,)
    assert get_origin(h29) is Literal and get_args(h29) == (2,)


def test_2025_input4_g10_h10_literal_pair_for_pv_base_denominator() -> None:
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    g10 = hints[format_key("Input 4 - External Financing", "G10")]
    h10 = hints[format_key("Input 4 - External Financing", "H10")]
    assert get_origin(g10) is Literal and get_args(g10) == (5,)
    assert get_origin(h10) is Literal and get_args(h10) == (10,)


