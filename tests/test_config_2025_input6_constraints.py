from __future__ import annotations

from typing import Literal, get_args, get_origin, get_type_hints

from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_input6_tailored_e40_template_yes_no_literals() -> None:
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    e40 = hints[format_key("Input 6 - Tailored Tests", "E40")]
    assert get_origin(e40) is Literal
    assert get_args(e40) == ("No",)


def test_2025_input6_optional_d18_d30_same_threshold_literal() -> None:
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    q6o = format_key("Input 6(optional)-Standard Test", "D18")
    d30 = format_key("Input 6(optional)-Standard Test", "D30")
    assert get_origin(hints[q6o]) is Literal
    assert hints[q6o] == hints[d30]
