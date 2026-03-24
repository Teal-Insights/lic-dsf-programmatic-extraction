from __future__ import annotations

from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
import re

import pytest

from src.lic_dsf_config import cells_in_range, normalize_cell_address


def _project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _validate_packaging():
    path = _project_root() / "scripts" / "validate_packaging_pyproject.py"
    spec = spec_from_file_location("validate_packaging_pyproject", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    mod = module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_pyproject_template_uses_modern_setuptools_backend() -> None:
    template = (_project_root() / "scripts" / "pyproject.toml.template").read_text(
        encoding="utf-8"
    )

    assert 'build-backend = "setuptools.build_meta"' in template
    assert "setuptools.backends._legacy:_Backend" not in template


def test_rendered_pyproject_template_passes_packaging_validation() -> None:
    template = (_project_root() / "scripts" / "pyproject.toml.template").read_text(
        encoding="utf-8"
    )
    rendered = template.replace("{{TEMPLATE_DATE}}", "2099-01-01").replace(
        "{{VERSION}}", "1.2.3"
    )
    mod = _validate_packaging()
    mod.validate_packaging_pyproject_text(rendered, source="rendered template")


def test_validate_packaging_pyproject_rejects_legacy_backend() -> None:
    sample = """
[project]
name = "x"
version = "0"
requires-python = ">=3.10"

[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.backends._legacy:_Backend"
"""
    mod = _validate_packaging()
    with pytest.raises(ValueError, match="build-backend"):
        mod.validate_packaging_pyproject_text(sample, source="fixture")


def test_deploy_script_does_not_track_python_sources_with_lfs() -> None:
    deploy_script = (_project_root() / "scripts" / "deploy.sh").read_text(
        encoding="utf-8"
    )

    assert (
        re.search(r"(?m)^\*\.py\s+filter=lfs\b", deploy_script) is None
    )


def test_deploy_script_validates_pyproject_after_render() -> None:
    deploy_script = (_project_root() / "scripts" / "deploy.sh").read_text(
        encoding="utf-8"
    )
    assert "validate_packaging_pyproject.py" in deploy_script


def test_normalize_cell_address_matches_cells_in_range_quoting() -> None:
    canonical = cells_in_range("Chart Data", "D10:D10")[0]
    assert normalize_cell_address("Chart Data!D10") == canonical
    assert normalize_cell_address("'Chart Data'!D10") == canonical
