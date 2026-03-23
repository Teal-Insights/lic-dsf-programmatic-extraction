"""Tests for indent-based hierarchical label detection."""

from __future__ import annotations

from unittest.mock import MagicMock

from src.lic_dsf_labels import (
    build_label_hierarchy,
    get_effective_indent,
    is_valid_label,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _mock_cell(value, indent: float = 0.0):
    """Create a mock cell with value and alignment.indent."""
    cell = MagicMock()
    cell.value = value
    alignment = MagicMock()
    alignment.indent = indent
    cell.alignment = alignment
    return cell


def _mock_ws(rows: dict[int, tuple[str | None, float]]) -> MagicMock:
    """Build a mock worksheet from {row: (value, indent)} for column 1.

    ``max_row`` is set to the largest row key.
    """
    ws = MagicMock()
    ws.max_row = max(rows.keys()) if rows else 1

    def cell(row: int, column: int) -> MagicMock:
        if row in rows:
            val, indent = rows[row]
            return _mock_cell(val, indent)
        return _mock_cell(None)

    ws.cell = cell
    return ws


# ---------------------------------------------------------------------------
# get_effective_indent
# ---------------------------------------------------------------------------

class TestGetEffectiveIndent:
    def test_no_indent_no_spaces(self):
        assert get_effective_indent(_mock_cell("Hello", 0.0)) == 0

    def test_alignment_indent_only(self):
        assert get_effective_indent(_mock_cell("Hello", 2.0)) == 2

    def test_leading_spaces_only(self):
        assert get_effective_indent(_mock_cell("   Hello", 0.0)) == 1

    def test_alignment_plus_leading_spaces(self):
        # alignment.indent=1 + leading spaces → 1 + 1 = 2
        assert get_effective_indent(_mock_cell("    Hello", 1.0)) == 2

    def test_none_value(self):
        assert get_effective_indent(_mock_cell(None, 0.0)) == 0

    def test_numeric_value(self):
        assert get_effective_indent(_mock_cell(42, 1.0)) == 1

    def test_no_alignment(self):
        cell = MagicMock()
        cell.value = "Hello"
        cell.alignment = None
        assert get_effective_indent(cell) == 0


# ---------------------------------------------------------------------------
# build_label_hierarchy – synthetic worksheets
# ---------------------------------------------------------------------------

class TestBuildLabelHierarchy:
    def test_flat_labels_no_ancestors(self):
        ws = _mock_ws({
            1: ("A", 0.0),
            2: ("B", 0.0),
            3: ("C", 0.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert h[1] == []
        assert h[2] == []
        assert h[3] == []

    def test_simple_parent_child(self):
        ws = _mock_ws({
            1: ("Parent", 0.0),
            2: ("Child", 1.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert h[1] == []
        assert h[2] == ["Parent"]

    def test_two_levels(self):
        ws = _mock_ws({
            1: ("Grandparent", 0.0),
            2: ("Parent", 1.0),
            3: ("Child", 2.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert h[1] == []
        assert h[2] == ["Grandparent"]
        assert h[3] == ["Grandparent", "Parent"]

    def test_sibling_resets_stack(self):
        ws = _mock_ws({
            1: ("Parent1", 0.0),
            2: ("Child1", 1.0),
            3: ("Parent2", 0.0),
            4: ("Child2", 1.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert h[2] == ["Parent1"]
        assert h[4] == ["Parent2"]

    def test_skip_empty_rows(self):
        ws = _mock_ws({
            1: ("Parent", 0.0),
            2: (None, 0.0),  # empty row
            3: ("Child", 1.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert 2 not in h  # skipped
        assert h[3] == ["Parent"]

    def test_invalid_labels_skipped(self):
        ws = _mock_ws({
            1: ("Parent", 0.0),
            2: ("...", 1.0),  # invalid label
            3: ("Child", 1.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert 2 not in h
        assert h[3] == ["Parent"]

    def test_placeholder_insert_filepath_skipped(self):
        ws = _mock_ws({
            1: ("[insert filepath]", 0.0),
            2: ("Real label", 1.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert 1 not in h  # placeholder filtered
        assert h[2] == []  # no ancestor because placeholder was skipped

    def test_pop_back_to_correct_level(self):
        # Parent > Child > Grandchild, then back to Parent > Child2
        ws = _mock_ws({
            1: ("Parent", 0.0),
            2: ("Child", 1.0),
            3: ("Grandchild", 2.0),
            4: ("Child2", 1.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert h[3] == ["Parent", "Child"]
        assert h[4] == ["Parent"]

    def test_leading_spaces_create_hierarchy(self):
        # Simulates the Paris Club pattern: indent=0 but leading spaces
        ws = _mock_ws({
            1: ("Official Bilaterals", 0.0),
            2: ("   Paris Club", 0.0),  # 3 leading spaces, indent=0
            3: ("Export Credit", 2.0),
        })
        h = build_label_hierarchy(ws, 1)
        assert h[1] == []
        assert h[2] == ["Official Bilaterals"]  # leading spaces → effective indent 1
        assert h[3] == ["Official Bilaterals", "Paris Club"]

    def test_min_max_row_bounds(self):
        ws = _mock_ws({
            1: ("Outside", 0.0),
            2: ("Parent", 0.0),
            3: ("Child", 1.0),
            4: ("Outside2", 0.0),
        })
        h = build_label_hierarchy(ws, 1, min_row=2, max_row=3)
        assert 1 not in h
        assert 4 not in h
        assert h[2] == []
        assert h[3] == ["Parent"]
