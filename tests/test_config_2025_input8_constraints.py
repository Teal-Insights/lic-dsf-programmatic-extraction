from __future__ import annotations

from typing import Literal, get_args, get_origin, get_type_hints

from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_input8_j27_s37_blank_leaves_for_offset_subgraph() -> None:
    """Non-formula blanks referenced under OFFSET/INDIRECT argument expansion."""
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    for coord in ("J27", "S37"):
        ann = hints[format_key("Input 8 - SDR", coord)]
        assert get_origin(ann) is Literal
        assert get_args(ann) == (None,)
