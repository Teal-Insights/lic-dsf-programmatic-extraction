from __future__ import annotations

from excel_grapher.grapher.dynamic_refs import format_key

from src.configs import load_template_config


def test_2025_config_defines_input5_dynamic_ref_constraints() -> None:
    """Input 5 cells referenced by OFFSET/INDIRECT must appear in LicDsfConstraints."""
    cfg = load_template_config("2025-08-12")
    ann = cfg.LicDsfConstraints.__annotations__
    samples = [
        format_key("Input 5 - Local-debt Financing", "C10"),
        format_key("Input 5 - Local-debt Financing", "AB63"),
        format_key("Input 5 - Local-debt Financing", "AE254"),
        format_key("Input 5 - Local-debt Financing", "AK302"),
        format_key("Input 5 - Local-debt Financing", "BU392"),
    ]
    for key in samples:
        assert key in ann, f"missing constraint for {key!r}"
