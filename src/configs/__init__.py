"""
Template configuration loader.

Each subdirectory under ``src/configs/`` corresponds to a template version
(named by date, e.g. ``2026-01-31``). This module provides utilities for
discovering available templates and loading their configuration.
"""

from __future__ import annotations

import importlib
import importlib.util
import types
from pathlib import Path

CONFIGS_DIR = Path(__file__).parent


def available_templates() -> list[str]:
    """Return sorted list of available template version names."""
    return sorted(
        d.name
        for d in CONFIGS_DIR.iterdir()
        if d.is_dir() and (d / "config.py").exists()
    )


def load_template_config(template: str) -> types.ModuleType:
    """
    Load and return the config module for the given template version.

    The returned module has attributes: WORKBOOK_PATH, WORKBOOK_TEMPLATE_URL,
    EXPORT_RANGES, REGION_CONFIG, PACKAGE_NAME, EXPORT_DIR,
    STRING_CONSTANT_EXCLUDES, BLANK_CONSTANT_EXCLUDES, get_dynamic_ref_config, etc.
    """
    config_path = CONFIGS_DIR / template / "config.py"
    if not config_path.exists():
        avail = available_templates()
        raise ValueError(
            f"Unknown template {template!r}. "
            f"Available: {', '.join(avail) if avail else '(none)'}"
        )
    package_path = f"src.configs.{template.replace('-', '_')}"
    try:
        return importlib.import_module(".config", package_path)
    except ModuleNotFoundError:
        module_name = f"src.configs.{template.replace('-', '_')}.config"
        spec = importlib.util.spec_from_file_location(module_name, config_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Cannot load config from {config_path}")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
