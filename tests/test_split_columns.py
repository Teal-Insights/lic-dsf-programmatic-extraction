from __future__ import annotations

from src.lic_dsf_export import split_columns_by_year_presence


def test_split_columns_by_year_presence_all_years() -> None:
    year_by_col: dict[int, int | None] = {1: 2020, 2: 2021, 3: 2022}
    assert split_columns_by_year_presence(start_col=1, end_col=3, year_by_col=year_by_col) == [
        ("year", 1, 3)
    ]


def test_split_columns_by_year_presence_all_meta() -> None:
    year_by_col: dict[int, int | None] = {1: None, 2: None}
    assert split_columns_by_year_presence(start_col=1, end_col=2, year_by_col=year_by_col) == [
        ("meta", 1, 2)
    ]


def test_split_columns_by_year_presence_mixed_segments() -> None:
    year_by_col = {1: None, 2: 2020, 3: 2021, 4: None, 5: 2022}
    assert split_columns_by_year_presence(start_col=1, end_col=5, year_by_col=year_by_col) == [
        ("meta", 1, 1),
        ("year", 2, 3),
        ("meta", 4, 4),
        ("year", 5, 5),
    ]
