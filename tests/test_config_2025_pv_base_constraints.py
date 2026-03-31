from __future__ import annotations

from typing import Literal, get_args, get_origin, get_type_hints

from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_pv_base_d198_d856_base_scalars_hundred() -> None:
    """Template leaves 100 on column D (not only AM:BP sweep rows)."""
    cfg = load_template_config("2025-08-12")
    hints = get_type_hints(cfg.LicDsfConstraints, include_extras=True)
    assert get_args(hints[format_key("PV_Base", "D40")]) == (3,)
    for coord in ("D198", "D648", "D856"):
        ann = hints[format_key("PV_Base", coord)]
        assert get_origin(ann) is Literal
        assert get_args(ann) == (100,)
