"""Tests that pipeline graph building enables safe identity-transit compression."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest


def test_build_graph_enables_provenance_and_compresses(monkeypatch: pytest.MonkeyPatch) -> None:
    mock_graph = MagicMock()
    mock_graph.compress_identity_transits = MagicMock()
    captured: dict = {}

    def fake_create(*args: object, **kwargs: object) -> MagicMock:
        captured["args"] = args
        captured["kwargs"] = kwargs
        return mock_graph

    monkeypatch.setattr("src.lic_dsf_pipeline.create_dependency_graph", fake_create)

    from src.lic_dsf_pipeline import build_graph

    wb = Path("dummy.xlsx")
    targets = ["Sheet1!A1"]
    g = build_graph(wb, targets, max_depth=7)

    assert g is mock_graph
    assert captured["kwargs"].get("capture_dependency_provenance") is True
    mock_graph.compress_identity_transits.assert_called_once_with()
