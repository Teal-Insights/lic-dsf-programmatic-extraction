"""Tests for scripts/profile_normalize_formula.py helpers."""

import importlib.util
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location(
    "profile_normalize_formula",
    _ROOT / "scripts" / "profile_normalize_formula.py",
)
assert _spec and _spec.loader
_mod = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _mod
_spec.loader.exec_module(_mod)


def test_truncate_short_unchanged() -> None:
    assert _mod._truncate("=A1+B1", 100) == "=A1+B1"


def test_truncate_newline_escaped() -> None:
    assert _mod._truncate("=A\n+B", 20) == "=A\\n+B"


def test_truncate_long_suffix() -> None:
    s = "=" + "X" * 200
    out = _mod._truncate(s, 20)
    assert out.endswith("...")
    assert len(out) == 20
