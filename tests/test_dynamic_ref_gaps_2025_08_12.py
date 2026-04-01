from __future__ import annotations

import pytest

from src.configs import load_template_config
from src.lic_dsf_pipeline import discover_targets, list_missing_dynamic_ref_leaves


@pytest.mark.slow
def test_2025_08_12_dynamic_ref_gaps_empty() -> None:
    """
    Ensure the 2025-08-12 template constraints cover all leaf candidates needed
    for OFFSET/INDIRECT/INDEX resolution, as reported by excel-grapher's
    list_dynamic_ref_constraint_candidates.

    This touches the real template workbook and is intentionally marked slow by
    repo-wide pytest defaults.
    """
    cfg = load_template_config("2025-08-12")
    targets = discover_targets(cfg.EXPORT_RANGES)
    gaps = list_missing_dynamic_ref_leaves(
        cfg.WORKBOOK_PATH,
        targets,
        max_depth=50,
        dynamic_refs=cfg.get_dynamic_ref_config(),
    )
    assert gaps == []


@pytest.mark.slow
def test_2025_08_12_dynamic_ref_candidates_are_constrained_no_infer() -> None:
    """
    Faster audit than `test_2025_08_12_dynamic_ref_gaps_empty`.

    This uses excel-grapher's `list_dynamic_ref_constraint_candidates` with
    `dynamic_refs=None`, which skips dynamic-target inference and returns the
    statically reachable leaf candidates feeding OFFSET/INDIRECT/INDEX arguments.

    We then assert every candidate has an annotation in our constraints type.
    """
    from excel_grapher.grapher import list_dynamic_ref_constraint_candidates

    cfg = load_template_config("2025-08-12")
    targets = discover_targets(cfg.EXPORT_RANGES)
    candidates = list_dynamic_ref_constraint_candidates(
        cfg.WORKBOOK_PATH,
        targets,
        dynamic_refs=None,
        max_depth=50,
    )
    hints = cfg.LicDsfConstraints.__annotations__
    missing = [c for c in candidates if c not in hints]
    assert missing == []

