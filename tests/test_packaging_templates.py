from __future__ import annotations

from pathlib import Path
import re


def _project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def test_pyproject_template_uses_modern_setuptools_backend() -> None:
    template = (_project_root() / "scripts" / "pyproject.toml.template").read_text(
        encoding="utf-8"
    )

    assert 'build-backend = "setuptools.build_meta"' in template
    assert "setuptools.backends._legacy:_Backend" not in template


def test_deploy_script_does_not_track_python_sources_with_lfs() -> None:
    deploy_script = (_project_root() / "scripts" / "deploy.sh").read_text(
        encoding="utf-8"
    )

    assert (
        re.search(r"(?m)^\*\.py\s+filter=lfs\b", deploy_script) is None
    )
