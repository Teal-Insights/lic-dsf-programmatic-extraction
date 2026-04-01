from __future__ import annotations

from typing import TypedDict

import pytest

from src.configs import load_template_config


def test_check_constraints_raises_when_required_specs_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    cfg = load_template_config("2025-08-12")

    class Empty(TypedDict, total=False):
        pass

    monkeypatch.setattr(
        cfg,
        "REQUIRED_CONSTRAINTS",
        ["'CI Summary'!B36", "'CI Summary'!Z999"],
    )
    with pytest.raises(ValueError, match="missing required"):
        cfg.check_constraints(Empty)


def test_check_constraints_ok_when_required_empty(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    cfg = load_template_config("2025-08-12")

    class Empty(TypedDict, total=False):
        pass

    monkeypatch.setattr(cfg, "REQUIRED_CONSTRAINTS", [])
    cfg.check_constraints(Empty)
