#!/usr/bin/env python3
"""Fail fast if a generated LIC-DSF package pyproject.toml has an invalid PEP 517 build backend."""

from __future__ import annotations

import argparse
import sys
import tomllib
from pathlib import Path

ALLOWED_BUILD_BACKENDS = frozenset({"setuptools.build_meta"})


def validate_packaging_pyproject_text(text: str, *, source: str = "pyproject.toml") -> None:
    try:
        data = tomllib.loads(text)
    except tomllib.TOMLDecodeError as exc:
        raise ValueError(f"{source}: invalid TOML ({exc})") from exc

    build = data.get("build-system")
    if not isinstance(build, dict):
        raise ValueError(f"{source}: missing or invalid [build-system] table")

    backend = build.get("build-backend")
    if backend not in ALLOWED_BUILD_BACKENDS:
        raise ValueError(
            f"{source}: [build-system] build-backend must be one of {sorted(ALLOWED_BUILD_BACKENDS)!r}, "
            f"got {backend!r}"
        )

    requires = build.get("requires")
    if not isinstance(requires, list) or not requires:
        raise ValueError(f"{source}: [build-system] requires must be a non-empty array")

    if not any(
        isinstance(item, str) and item.lower().startswith("setuptools") for item in requires
    ):
        raise ValueError(
            f"{source}: [build-system] requires must include a setuptools requirement, got {requires!r}"
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "path",
        type=Path,
        help="Path to pyproject.toml",
    )
    args = parser.parse_args(argv)
    text = args.path.read_text(encoding="utf-8")
    try:
        validate_packaging_pyproject_text(text, source=str(args.path))
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
