from __future__ import annotations

import lic_dsf_pipeline as pipeline


class _DummyGraph:
    pass


class _DummyGenerator:
    def __init__(self, graph: _DummyGraph) -> None:
        self.graph = graph

    def classify_leaf_nodes(
        self,
        targets,
        constant_types=None,
        constant_ranges=None,
        constant_blanks=False,
        attach_to_graph=False,
    ):
        return ["Sheet1!A1", "Sheet1!B2"], ["Sheet1!C3"]


def test_classify_input_addresses_returns_set(monkeypatch) -> None:
    monkeypatch.setattr(pipeline, "CodeGenerator", _DummyGenerator)

    inputs = pipeline.classify_input_addresses(
        _DummyGraph(),
        ["Sheet1!Z9"],
        constant_ranges=["Sheet1!X1:X2"],
        constant_blanks=True,
    )

    assert inputs == {"Sheet1!A1", "Sheet1!B2"}
