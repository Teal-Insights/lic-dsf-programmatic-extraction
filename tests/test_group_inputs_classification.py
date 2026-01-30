from __future__ import annotations

from dataclasses import dataclass

import lic_dsf_pipeline as pipeline


@dataclass(frozen=True)
class _DummyNode:
    sheet: str
    address: str
    formula: str | None
    is_leaf: bool
    value: object


class _DummyGraph:
    def __init__(self, nodes: dict[str, _DummyNode] | None = None) -> None:
        self._nodes = nodes or {}

    def __iter__(self):
        return iter(self._nodes.keys())

    def get_node(self, key: str):
        return self._nodes.get(key)


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


def test_classify_input_addresses_readds_blank_excludes(monkeypatch) -> None:
    nodes = {
        "blank": _DummyNode(
            "Input 6(optional)-Standard Test", "D8", None, True, None
        ),
    }
    graph = _DummyGraph(nodes)
    monkeypatch.setattr(pipeline, "CodeGenerator", _DummyGenerator)

    inputs = pipeline.classify_input_addresses(
        graph,
        ["Sheet1!Z9"],
        constant_blanks=True,
        blank_excludes={"'Input 6(optional)-Standard Test'!D8"},
    )

    assert "'Input 6(optional)-Standard Test'!D8" in inputs
