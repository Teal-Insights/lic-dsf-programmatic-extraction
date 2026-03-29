"""Tests for formula-based year-offset detection in header rows."""

from __future__ import annotations

from typing import Mapping
from unittest.mock import MagicMock

from src.configs import load_template_config

_cfg = load_template_config("2026-01-31")
WORKBOOK_PATH = _cfg.WORKBOOK_PATH
import pytest
from src.lic_dsf_labels import (
    detect_year_offset_headers,
)


def _mock_ws(
    formulas: Mapping[tuple[int, int], str | None],
    values: Mapping[tuple[int, int], object] | None = None,
    *,
    force_empty_cells: bool = False,
) -> tuple[MagicMock, MagicMock]:
    """Build a (ws_formulas, ws_values) mock pair."""
    values = values or {}

    def _make(cell_map: Mapping[tuple[int, int], object]) -> MagicMock:
        ws = MagicMock()

        def cell(row: int, column: int) -> MagicMock:
            c = MagicMock()
            c.value = cell_map.get((row, column))
            return c

        ws.cell = cell
        ws.max_column = max((c for r, c in cell_map), default=1)
        ws._cells = {} if force_empty_cells else dict(cell_map)
        return ws

    return _make(formulas), _make(values)


# --- anchor detection ---


def test_detects_projection_year_anchor() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={(8, 4): "=ProjectionYear", (8, 5): "=D8+1"},
        values={(8, 4): 0, (8, 5): 1},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 8)
    assert result[4] == 0
    assert result[5] == 1


def test_detects_macro_debt_data_u4_anchor() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={(8, 4): "='Macro-Debt_Data'!U4", (8, 5): "=D8+1"},
        values={(8, 4): 0, (8, 5): 1},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 8)
    assert result[4] == 0
    assert result[5] == 1


def test_detects_macro_debt_data_u5_anchor() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={(3, 4): "='Macro-Debt_Data'!U5", (3, 5): "=D3+1"},
        values={(3, 4): 0, (3, 5): 1},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 3)
    assert result[4] == 0
    assert result[5] == 1


# --- chain walking ---


def test_walks_plus_one_chain_right() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={
            (8, 4): "=ProjectionYear",
            (8, 5): "=D8+1",
            (8, 6): "=E8+1",
            (8, 7): "=F8+1",
        },
        values={(8, 4): 0, (8, 5): 1, (8, 6): 2, (8, 7): 3},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 8)
    assert result == {4: 0, 5: 1, 6: 2, 7: 3}


def test_walks_minus_one_chain_left() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={
            (8, 3): "=D8-1",
            (8, 4): "=ProjectionYear",
            (8, 5): "=D8+1",
        },
        values={(8, 3): -1, (8, 4): 0, (8, 5): 1},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 8)
    assert result == {3: -1, 4: 0, 5: 1}


# --- indirect references (row copies another row) ---


def test_detects_indirect_row_reference() -> None:
    """Ext_Debt_Data row 1 copies from row 9 with =+E9, =+F9, etc."""
    ws_f, ws_v = _mock_ws(
        formulas={(1, 5): "=+E9", (1, 6): "=+F9", (1, 7): "=+G9"},
        values={(1, 5): -1, (1, 6): 0, (1, 7): 1},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 1)
    assert result == {5: -1, 6: 0, 7: 1}


# --- skips non-year-offset formulas ---


def test_ignores_translation_formulas() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={
            (5, 2): '=++IF(START!L10="English",translation!C218,"")',
            (5, 4): "=ProjectionYear",
            (5, 5): "=D5+1",
        },
        values={(5, 2): "Further Details", (5, 4): 0, (5, 5): 1},
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 5)
    assert 2 not in result
    assert result[4] == 0
    assert result[5] == 1


def test_empty_row_returns_empty() -> None:
    ws_f, ws_v = _mock_ws(formulas={}, values={})
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 8)
    assert result == {}


def test_empty_cells_fall_back_to_max_column_scan() -> None:
    ws_f, ws_v = _mock_ws(
        formulas={(8, 4): "=ProjectionYear", (8, 5): "=D8+1"},
        values={(8, 4): 0, (8, 5): 1},
        force_empty_cells=True,
    )
    result = detect_year_offset_headers(ws_f, ws_v, "TestSheet", 8)
    assert result[4] == 0
    assert result[5] == 1


# --- integration with real workbook ---


@pytest.mark.slow
def test_b1_gdp_ext_row8_offsets() -> None:
    """B1_GDP_ext header row 8 should yield offsets from -1 through 20."""
    import fastpyxl

    wb_f = fastpyxl.load_workbook(WORKBOOK_PATH)
    wb_v = fastpyxl.load_workbook(WORKBOOK_PATH, data_only=True)
    try:
        result = detect_year_offset_headers(
            wb_f["B1_GDP_ext"], wb_v["B1_GDP_ext"], "B1_GDP_ext", 8
        )
        assert len(result) >= 20
        offsets = sorted(result.values())
        assert offsets[0] == -1
        assert offsets[-1] == 20
        # offset 0 should be present (the anchor)
        assert 0 in result.values()
    finally:
        wb_f.close()
        wb_v.close()
