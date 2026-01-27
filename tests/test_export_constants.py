from __future__ import annotations

from pathlib import Path

from dataclasses import dataclass

from lic_dsf_pipeline import iter_string_constant_addresses


@dataclass(frozen=True)
class _DummyNode:
    sheet: str
    address: str
    formula: str | None
    is_leaf: bool
    value: object


class _DummyGraph:
    def __init__(self, nodes: dict[str, _DummyNode]) -> None:
        self._nodes = nodes

    def __iter__(self):
        return iter(self._nodes.keys())

    def get_node(self, key: str):
        return self._nodes.get(key)


def test_iter_string_constant_addresses_excludes_and_quotes() -> None:
    nodes = {
        "a": _DummyNode("Sheet1", "A1", None, True, "hello"),
        "b": _DummyNode("Sheet1", "A2", None, True, 10),
        "c": _DummyNode("Sheet1", "A3", "=A1", False, "ignored"),
        "d": _DummyNode("Weird Sheet", "B2", None, True, "value"),
    }
    graph = _DummyGraph(nodes)

    ranges = iter_string_constant_addresses(graph, {"Sheet1!A1"})

    assert "Sheet1!A1" not in ranges
    assert "'Weird Sheet'!B2" in ranges
