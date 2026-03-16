"""
Template-specific configuration for LIC-DSF template 2026-01-31.

This module contains all configuration that is specific to this template version:
workbook path, export ranges, region config, constraints, and constant excludes.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal, cast, TypedDict

from excel_grapher.grapher import DynamicRefConfig

from ...lic_dsf_config import ExportRangeConfig
from ...lic_dsf_labels import RegionConfig


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

WORKBOOK_PATH = Path("workbooks/lic-dsf-template-2025-08-12.xlsm")
WORKBOOK_TEMPLATE_URL = (
    "https://thedocs.worldbank.org/en/doc/f0ade6bcf85b6f98dbeb2c39a2b7770c-0360012025/original/LIC-DSF-IDA21-Template-08-12-2025-vf.xlsm"
)

# ---------------------------------------------------------------------------
# Export package
# ---------------------------------------------------------------------------

PACKAGE_NAME = "lic_dsf_2025_08_12"
EXPORT_DIR = Path("dist/lic-dsf-2025-08-12")

# ---------------------------------------------------------------------------
# Export ranges
# ---------------------------------------------------------------------------

STRESS_TEST_ROW_LABELS: list[str] = [
    "Baseline",
    "A1. Key variables at their historical averages in 2024-2034 2/",
    "B1. Real GDP growth",
    "B2. Primary balance",
    "B3. Exports",
    "B4. Other flows 3/",
    "B5. Depreciation",
    "B6. Combination of B1-B5",
    "",
    "C1. Combined contingent liabilities",
    "C2. Natural disaster",
    "C3. Commodity price",
    "C4. Market Financing",
    "A2. Alternative Scenario :[Customize, enter title]",
]


STRESS_TEST_BLOCKS: list[tuple[str, int]] = [
    ("PV of Debt-to-GDP Ratio", 239),
    ("PV of Debt-to-Revenue Ratio", 281),
    ("Debt Service-to-Revenue Ratio", 318),
    ("Debt Service-to-GDP Ratio", 351),
]


EXPORT_FIXED_RANGES: list[ExportRangeConfig] = [
    {
        "label": "External DSA risk rating signals",
        "range_spec": "'Chart Data'!D10:D17",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Fiscal (Total Public Debt) risk rating signals",
        "range_spec": "'Chart Data'!I10:I14",
        "entrypoint_mode": "per_cell",
    },
    {
        "label": "Applicable tailored stress test signals",
        "range_spec": "'Chart Data'!I17:I19",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Fiscal space for moderate risk category",
        "range_spec": "'Chart Data'!E25:E27",
        "entrypoint_mode": "row_group",
    },
    {
        "label": "Overall rating",
        "range_spec": "'Chart Data'!L10:L11",
        "entrypoint_mode": "row_group",
    },
]


def _export_chart_data_ranges() -> list[ExportRangeConfig]:
    out: list[ExportRangeConfig] = list(EXPORT_FIXED_RANGES)
    for metric_label, start_row in STRESS_TEST_BLOCKS:
        for i, row_label in enumerate(STRESS_TEST_ROW_LABELS):
            if not row_label:
                continue
            row = start_row + i
            out.append(
                {
                    "label": f"{metric_label} - {row_label}",
                    "range_spec": f"'Chart Data'!D{row}:X{row}",
                    "entrypoint_mode": "row_group",
                }
            )
    return out


EXPORT_RANGES: list[ExportRangeConfig] = _export_chart_data_ranges()

# ---------------------------------------------------------------------------
# Region config (label extraction)
# ---------------------------------------------------------------------------

REGION_CONFIG: list[RegionConfig] = [
    {
        "sheet": "Input 5 - Local-debt Financing",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [5],
        "label_columns": ["A", "B"],
    },
    {
        "sheet": "Ext_Debt_Data",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [1, 9],
        "label_columns": ["A"],
    },
    {
        "sheet": "PV_Base",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [7],
        "label_columns": ["A", "C"],
    },
    {
        "sheet": "PV_LC_NR1",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [3],
        "label_columns": ["A", "C"],
    },
    {
        "sheet": "PV_LC_NR2",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [3],
        "label_columns": ["A", "C"],
    },
    {
        "sheet": "PV_LC_NR3",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [3],
        "label_columns": ["A", "C"],
    },
    {
        "sheet": "Input 3 - Macro-Debt data(DMX)",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [7],
        "label_columns": ["A", "B", "C"],
    },
    {
        "sheet": "Input 4 - External Financing",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [6],
        "label_columns": ["B"],
    },
    {
        "sheet": "Baseline - external",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [8],
        "label_columns": ["B"],
    },
    {
        "sheet": "Baseline - public",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [7],
        "label_columns": ["B"],
    },
    {
        "sheet": "Input 8 - SDR",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [9],
        "label_columns": ["A"],
    },
    {
        "sheet": "B1_GDP_ext",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [8],
        "label_columns": ["B", "Z"],
    },
    {
        "sheet": "B3_Exports_ext",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [8],
        "label_columns": ["B", "Z"],
    },
    {
        "sheet": "B4_other flows_ext",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [8],
        "label_columns": ["B", "Z"],
    },
    {
        "sheet": "Macro-Debt_Data",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [1, 5],
        "label_columns": ["B", "E"],
    },
    {
        "sheet": "PV Stress",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [3],
        "label_columns": ["A", "C"],
    },
    {
        "sheet": "Input 6(optional)-Standard Test",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [],
        "label_columns": ["A", "B", "C"],
    },
    {
        "sheet": "Input 7 - Residual Financing",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [],
        "label_columns": ["B"],
    },
    {
        "sheet": "START",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [],
        "label_columns": ["A", "B"],
    },
    {
        "sheet": "lookup",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [],
        "label_columns": ["A", "B"],
    },
    {
        "sheet": "translation",
        "min_row": None,
        "max_row": None,
        "min_col": None,
        "max_col": None,
        "header_rows": [],
        "label_columns": ["A", "B"],
    },
]

# ---------------------------------------------------------------------------
# Constraints (OFFSET / INDIRECT resolution)
# ---------------------------------------------------------------------------

"""
Dynamic refs (OFFSET/INDIRECT) are resolved via a constraint-based config.
Iterative workflow: run the script; if DynamicRefError is raised, the message
includes the formula cell that needs a constraint. Inspect that cell and the
row/column headers in the workbook to decide plausible input domains, add the
address to LicDsfConstraints (with Annotated[int, Between(lo, hi)] or
Literal[...]), then re-run until the graph builds.
"""


class LicDsfConstraints(TypedDict, total=False):
    pass


LiteralType = cast(Any, Literal)

# Language selector and lookup table.
_LANG = Literal["English", "French", "Portuguese", "Spanish"]
_LANG_LOOKUP = Literal[
    "English", "French", "Portuguese", "Spanish", "Français", "Portugues", "Español"
]
LicDsfConstraints.__annotations__["START!L10"] = _LANG
LicDsfConstraints.__annotations__["START!K10"] = _LANG
for _r in range(4, 8):
    for _c in ("BB", "BC"):
        LicDsfConstraints.__annotations__[f"lookup!{_c}{_r}"] = _LANG_LOOKUP


def _constrain_lookup_countries(constraints: type[Any]) -> None:
    _countries: list[tuple[int, str]] = [
        (4, 'Afghanistan'),
        (5, 'Bangladesh'),
        (6, 'Benin'),
        (7, 'Bhutan'),
        (8, 'Burkina Faso'),
        (9, 'Burundi'),
        (10, 'Cambodia'),
        (11, 'Cameroon'),
        (12, 'Cabo Verde'),
        (13, 'Central African Republic'),
        (14, 'Chad'),
        (15, 'Comoros'),
        (16, 'Congo, DR'),
        (17, 'Congo, Republic of'),
        (18, "Cote d'Ivoire"),
        (19, 'Djibouti'),
        (20, 'Dominica'),
        (21, 'Eritrea'),
        (22, 'Ethiopia'),
        (23, 'Gambia, The'),
        (24, 'Ghana'),
        (25, 'Grenada'),
        (26, 'Guinea'),
        (27, 'Guinea-Bissau'),
        (28, 'Guyana'),
        (29, 'Haiti'),
        (30, 'Honduras'),
        (31, 'Kenya'),
        (32, 'Kiribati'),
        (33, 'Kyrgyz Republic'),
        (34, 'Lao PDR'),
        (35, 'Lesotho'),
        (36, 'Liberia'),
        (37, 'Madagascar'),
        (38, 'Malawi'),
        (39, 'Maldives'),
        (40, 'Mali'),
        (41, 'Marshall Islands'),
        (42, 'Mauritania'),
        (43, 'Micronesia'),
        (44, 'Moldova'),
        (45, 'Mozambique'),
        (46, 'Myanmar'),
        (47, 'Nepal'),
        (48, 'Nicaragua'),
        (49, 'Niger'),
        (50, 'Papua New Guinea'),
        (51, 'Rwanda'),
        (52, 'Samoa'),
        (53, 'Sao Tome & Principe'),
        (54, 'Senegal'),
        (55, 'Sierra Leone'),
        (56, 'Solomon Islands'),
        (57, 'Somalia'),
        (58, 'South Sudan'),
        (59, 'St. Lucia'),
        (60, 'St. Vincent & the Grenadines'),
        (61, 'Sudan'),
        (62, 'Tajikistan'),
        (63, 'Tanzania'),
        (64, 'Timor-Leste'),
        (65, 'Togo'),
        (66, 'Tonga'),
        (67, 'Tuvalu'),
        (68, 'Uganda'),
        (69, 'Uzbekistan'),
        (70, 'Vanuatu'),
        (71, 'Yemen, Republic of'),
        (72, 'Zambia'),
        (73, 'Zimbabwe'),
    ]
    for _row, _name in _countries:
        constraints.__annotations__[f"lookup!C{_row}"] = LiteralType[_name]


_constrain_lookup_countries(LicDsfConstraints)


def get_dynamic_ref_config() -> DynamicRefConfig:
    """Return a DynamicRefConfig for constraint-based resolution of OFFSET/INDIRECT."""
    return DynamicRefConfig.from_constraints_and_workbook(
        LicDsfConstraints, WORKBOOK_PATH
    )


# ---------------------------------------------------------------------------
# Constant excludes (for input classification)
# ---------------------------------------------------------------------------

STRING_CONSTANT_EXCLUDES = {
    "START!K10",
    "'BLEND floating calculations WB'!C6",
    "'Input 6(optional)-Standard Test'!C4",
    "'Input 6(optional)-Standard Test'!C5",
    "'Input 6(optional)-Standard Test'!C7",
    "'Input 6(optional)-Standard Test'!C8",
    "'Input 6(optional)-Standard Test'!D18",
    "'Input 6(optional)-Standard Test'!D26",
    "'Input 6(optional)-Standard Test'!D30",
    "'Input 6(optional)-Standard Test'!D33",
    "'Input 6(optional)-Standard Test'!D8",
    "'Input 6(optional)-Standard Test'!D9",
}
BLANK_CONSTANT_EXCLUDES = {
    "'Input 6(optional)-Standard Test'!D8",
    "'Input 6(optional)-Standard Test'!D9",
}
