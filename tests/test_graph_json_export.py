"""Tests for dependency graph JSON export."""

from __future__ import annotations

import json
from pathlib import Path

from excel_grapher.grapher import DependencyGraph, Node

from src.lic_dsf_pipeline import dependency_graph_to_dict, export_graph_json


def test_dependency_graph_to_dict_structure() -> None:
    g = DependencyGraph()
    n_leaf = Node("Data", "A", 1, None, None, 3.14, True, {"row_labels": ["GDP"]})
    n_f = Node("Data", "B", 1, "=A1*2", "=A1*2", None, False, {})
    g.add_node(n_leaf)
    g.add_node(n_f)
    g.add_edge(n_f.key, n_leaf.key)

    d = dependency_graph_to_dict(g)
    assert d["version"] == 1
    assert set(d["nodes"]) == {n_leaf.key, n_f.key}
    assert d["nodes"][n_f.key]["formula"] == "=A1*2"
    assert d["nodes"][n_leaf.key]["metadata"]["row_labels"] == ["GDP"]
    edges = {(e["from"], e["to"]) for e in d["edges"]}
    assert (n_f.key, n_leaf.key) in edges


def test_export_graph_json_writes_valid_json(tmp_path: Path) -> None:
    g = DependencyGraph()
    n = Node("S", "C", 5, None, None, None, True, {})
    g.add_node(n)
    path = tmp_path / "graph.json"
    export_graph_json(g, path)
    loaded = json.loads(path.read_text(encoding="utf-8"))
    assert loaded["version"] == 1
    assert n.key in loaded["nodes"]
    assert loaded["nodes"][n.key]["sheet"] == "S"
    assert loaded["nodes"][n.key]["row"] == 5
