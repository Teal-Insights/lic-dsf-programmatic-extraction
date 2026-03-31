from __future__ import annotations

from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_config_defines_input5_dynamic_ref_constraints() -> None:
    """Spot-check Input 5 leaves that remain in LicDsfConstraints (wide AG:BT grids are on-demand)."""
    cfg = load_template_config("2025-08-12")
    ann = cfg.LicDsfConstraints.__annotations__
    samples = [
        format_key("Input 5 - Local-debt Financing", "C10"),
        format_key("Input 5 - Local-debt Financing", "AB63"),
        format_key("Input 5 - Local-debt Financing", "H254"),
        format_key("Input 5 - Local-debt Financing", "M466"),
        format_key("Input 5 - Local-debt Financing", "AF250"),
    ]
    for key in samples:
        assert key in ann, f"missing constraint for {key!r}"
