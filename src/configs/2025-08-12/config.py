"""
Template-specific configuration for LIC-DSF template 2025-08-12.

This module contains all configuration that is specific to this template version:
workbook path, export ranges, region config, constraints, and constant excludes.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Literal, TypedDict, Annotated, cast, get_type_hints

import fastpyxl
from fastpyxl.utils.cell import column_index_from_string, range_boundaries, get_column_letter
from fastpyxl.worksheet.formula import ArrayFormula

from excel_grapher import RealBetween, constrain
from excel_grapher.grapher import DynamicRefConfig
from excel_grapher.grapher.dynamic_refs import format_key
from excel_grapher.core.cell_types import Between

from src.lic_dsf_config import (
    ExportRangeConfig,
    WorkbookMetadata,
    cells_in_range,
    parse_range_spec,
)
from src.lic_dsf_labels import RegionConfig


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

WORKBOOK_PATH = Path("workbooks/lic-dsf-template-2025-08-12.xlsm")

# Dependency graph / INDIRECT sometimes uses the codename `Market_financing`; the workbook tab is
# `C4_Market_financing`. Leaf verification resolves this map to the physical sheet.
_CONSTRAINT_VERIFY_SHEET_ALIASES: dict[str, str] = {"Market_financing": "C4_Market_financing"}
WORKBOOK_TEMPLATE_URL = (
    "https://thedocs.worldbank.org/en/doc/f0ade6bcf85b6f98dbeb2c39a2b7770c-0360012025/original/LIC-DSF-IDA21-Template-08-12-2025-vf.xlsm"
)
WORKBOOK_METADATA: WorkbookMetadata = {
    "creator": "spalazzo",
    "created": "2002-02-01",
    "modified": "2025-09-18T22:03:17",
}

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


FIGURE_DATA_ROWS: list[int] = [
    # Figure 1 (Output 2-1 Stress_Charts_Ex)
    51,
    61,
    62,
    63,
    64,
    66,
    93,
    103,
    104,
    105,
    106,
    108,
    135,
    145,
    146,
    147,
    148,
    150,
    177,
    187,
    188,
    189,
    190,
    192,
    # Figure 2 extras (Output 2-2 Stress_Charts_Pub)
    263,
    264,
    265,
    267,
    306,
    341,
    342,
    343,
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
    seen_row_specs = {entry["range_spec"] for entry in out}

    def add_chart_data_row(row: int, label: str) -> None:
        range_spec = f"'Chart Data'!D{row}:X{row}"
        if range_spec in seen_row_specs:
            return
        out.append(
            {
                "label": label,
                "range_spec": range_spec,
                "entrypoint_mode": "row_group",
            }
        )
        seen_row_specs.add(range_spec)

    for metric_label, start_row in STRESS_TEST_BLOCKS:
        for i, row_label in enumerate(STRESS_TEST_ROW_LABELS):
            if not row_label:
                continue
            row = start_row + i
            add_chart_data_row(row, f"{metric_label} - {row_label}")

    for row in FIGURE_DATA_ROWS:
        add_chart_data_row(row, f"Figure data row {row}")

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
Dynamic refs (OFFSET/INDIRECT/INDEX) are resolved via a constraint-based config. Iterative workflow: run the script; if DynamicRefError is raised, the message includes the formula cell whose inputs need constraints. Inspect that cell and the row/column headers in the workbook to decide plausible input domains, use `constrain` to set a constraint (e.g., `constrain(LicDsfConstraints, "'Sheet Name'!A965", Literal["value"])` for a constant or `Annotated[float, RealBetween(min=0)]` for a real-valued numeric range; use `Annotated[int, Between(lo, hi)]` for discrete integer bounds).

Then re-run until the graph builds. Note that if row/column labels or intentionally blank cells show up in error output, they have been referenced by a dynamic ref and must be constrained for the graph to resolve. Blank cells can be set to `Literal[None]`.

The goal is to set sensible constraints that reflect the range of sane values we will allow for the cells. To determine the plausible range of input values, investigate the cells by using enrichment_audit.json (or the heuristic label-scanning tools in src/lic_dsf_labels.py) to see their labels, and fastpyxl to check their current values. In addition to the empty template workbook, workbooks/lic-dsf-template-2025-08-12.xlsm, we also have one filled out with illustrative data: workbooks/dsf-uga.xlsm. This may be helpful for identifying which cells are structurally blank (will be blank in the demo workbook) as opposed to template input blanks meant to be filled in by users.

When the template workbook is present, ``verify_lic_dsf_constraints_target_leaves`` scans every constrained cell in the template and raises if any contain an Excel formula. Constraints are applied only to leaves.

Please include comments to explain decisions about plausible cell domains in terms of what they represent in the workbook.
"""

LiteralType = cast(Any, Literal)

# Constraint types for cells that feed OFFSET/INDIRECT. Keys must be address-style (e.g. "Sheet1!B1").
# Add entries when the script raises DynamicRefError: the message lists leaf cells that need
# constraints. Add each to __annotations__ (with Annotated[int, Between(lo, hi)],
# Annotated[float, RealBetween(...)], Annotated[..., FromWorkbook()], or Literal[...])
# then re-run. Repeat until the graph builds.
class LicDsfConstraints(TypedDict, total=False):
    pass

# Lookup switches; treat as constants
constrain(LicDsfConstraints, "lookup!AF4", Literal["New"])
constrain(LicDsfConstraints, "lookup!AF5", Literal["Old"])

# Marker to use for applicable tailored stress test; we can treat as a constant
constrain(LicDsfConstraints, "'Chart Data'!I21", Literal[1])

# Year header slot on row 35 (empty in template; W35/X35 are 2043/2044); feeds Chart Data dynamic refs.
constrain(LicDsfConstraints, "'Chart Data'!Y35", Annotated[int | None, Between(1990, 2100)])

# PV_Base!B9xx = CONCAT("$", A9xx, "$", $A$<row>) → INDIRECT($B9xx). Row-index cells A917, A941, A965 (fixed).
# Treat these as constants derived from the current workbook values.
constrain(LicDsfConstraints, "PV_Base!A917", Literal[64])
constrain(LicDsfConstraints, "PV_Base!A941", Literal[90])
constrain(LicDsfConstraints, "PV_Base!A965", Literal[115])

constrain(LicDsfConstraints, "PV_Base!A965", Annotated[float, RealBetween(min=0)])

# A918:A938, A942:A962, A966:A986 each has a single cached letter D, E, …, X.
for _start, _end in [(918, 939), (942, 963), (966, 987)]:
    for _row in range(_start, _end):
        _letter = chr(ord("D") + _row - _start)
        LicDsfConstraints.__annotations__[f"PV_Base!A{_row}"] = LiteralType[_letter]

# Language selector and lookup table (feed INDIRECT/VLOOKUP for language-dependent refs).
# START!L10 = VLOOKUP(K10, lookup!BB4:BC7, 2) — formula cell; only K10 is a leaf input.
# Each lookup row must be Literal[...] for that cell only: a shared 7-value union on the whole
# BB4:BC7 range makes the engine enumerate 7^8 combinations for INDIRECT fallbacks.
_LANG = Literal["English", "French", "Portuguese", "Spanish"]
constrain(LicDsfConstraints, "START!K10", _LANG)
constrain(LicDsfConstraints, "lookup!BB4", Literal["English"])
constrain(LicDsfConstraints, "lookup!BC4", Literal["English"])
constrain(LicDsfConstraints, "lookup!BB5", Literal["Français"])
constrain(LicDsfConstraints, "lookup!BC5", Literal["French"])
constrain(LicDsfConstraints, "lookup!BB6", Literal["Portugues"])
constrain(LicDsfConstraints, "lookup!BC6", Literal["Portuguese"])
constrain(LicDsfConstraints, "lookup!BB7", Literal["Español"])
constrain(LicDsfConstraints, "lookup!BC7", Literal["Spanish"])


# ---------------------------------------------------------------------------
# Market financing (C4 stress test sheet)
# ---------------------------------------------------------------------------

# C4_Market_financing holds the tailored “C4. Market Financing” stress scenario: layout and
# label cells, optional structural blanks, and user parameters in AB (shock toggles, haircuts,
# rate spreads). Domains for true leaves are applied in `_apply_lic_dsf_workbook_leaf_overlays`
# so formula rows are skipped per cell.


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
    constrain(LicDsfConstraints, f"lookup!C{_row}", LiteralType[_name])


def _constrain_pv_baseline_com(constraints: type[Any]) -> None:
    # Non-negative financial flows / values (or None for empty cells)
    financial_type = Annotated[float | None, RealBetween(0, 1e15)]

    # B column mirrors Input 4 G/H via formulas — constrain Input 4 leaves, not PV_baseline_com!B*.

    # D column: mixed constants and financial values
    # D7: total commercial (financial)
    constrain(constraints, "PV_baseline_com!D7", financial_type)
    # D19, D45, D71, D97, D123: "Base" normalization = 100
    for r in (19, 45, 71, 97, 123):
        constrain(constraints, f"PV_baseline_com!D{r}", Literal[100])
    # D32, D58, D84, D110, D136: "New forex borrowing (gross, USD)"
    for r in (32, 58, 84, 110, 136):
        constrain(constraints, f"PV_baseline_com!D{r}", financial_type)

    # AF33, AF59, AF85, AF111, AF137: "cumulative" in Output sections (Eurobond thru COM5)
    for r in (33, 59, 85, 111, 137):
        constrain(constraints, f"PV_baseline_com!AF{r}", financial_type)

    # BD23, BD49, BD75, BD101, BD127: "Total debt service"
    for r in (23, 49, 75, 101, 127):
        constrain(constraints, f"PV_baseline_com!BD{r}", financial_type)
        constrain(constraints, f"PV_baseline_com!AR{r}:BP{r}", financial_type)

    # H:AE ranges for "New forex borrowing (gross, USD)" rows
    cols = (
        "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
        "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE"
    )
    for r in (32, 58, 84, 110, 136):
        for col in cols:
            constrain(constraints, f"PV_baseline_com!{col}{r}", financial_type)
        constrain(constraints, f"PV_baseline_com!AR{r}:BP{r}", financial_type)


_constrain_pv_baseline_com(LicDsfConstraints)


def _constrain_pv_stress_and_pv_base_index_cells(constraints: type[Any]) -> None:
    """INDEX/OFFSET inputs on PV Stress and PV_Base (labels from enrichment_audit.json).

    PV Stress: interest and USD discount columns → unit rates; borrowing and cumulative → flows.
    PV_Base AF: cumulative outputs; BD: total debt service; D: Interest rates, Base=100 scalars,
    IDA line, or maturity/Base blocks.
    """
    financial_type = Annotated[float | None, RealBetween(0, 1e15)]
    unit_rate = Annotated[float | None, RealBetween(0, 1)]

    # Creditor label column: static text leaves (same strings as Input 4 B67:B75), not G/H formulas.
    for _br, _bv in (
        (35, "IDA - small economy"),
        (36, "IDA - regular"),
        (37, "IDA - blend"),
        (38, "IDA - SML"),
        (39, "IDA - 50Y loans"),
        (41, "IDA NEW 40-year credits"),
        (42, "IDA NEW Regular"),
        (43, "IDA NEW Blend (also enter) -->"),
        (44, "IDA NEW 60-year credits"),
    ):
        constrain(constraints, f"PV_Base!B{_br}", LiteralType[_bv])  # ty: ignore[invalid-type-form]

    constrain(constraints, "'PV Stress'!D147", unit_rate)
    constrain(constraints, "'PV Stress'!D161", financial_type)
    constrain(constraints, "'PV Stress'!D4", financial_type)
    constrain(constraints, "'PV Stress'!E161:G161", financial_type)
    constrain(constraints, "'PV Stress'!H147:X147", unit_rate)
    constrain(constraints, "'PV Stress'!Y148:AF148", unit_rate)
    constrain(constraints, "'PV Stress'!Y162:AF162", financial_type)
    constrain(constraints, "'PV Stress'!Y30:AF30", financial_type)

    for _r in (
        23,
        116,
        246,
        272,
        298,
        324,
        350,
        376,
        402,
        428,
        454,
        480,
        506,
        532,
        558,
        584,
        610,
        636,
        662,
        688,
        714,
        740,
        766,
        792,
        818,
        844,
        870,
        896,
    ):
        constrain(constraints, f"PV_Base!AF{_r}", financial_type)

    for _r in (366, 470, 496, 600, 626, 730, 756, 808, 834, 886):
        constrain(constraints, f"PV_Base!BD{_r}", financial_type)

    for _r in (
        27,
        276,
        302,
        328,
        354,
        380,
        406,
        432,
        458,
        484,
        510,
        536,
        562,
        588,
        614,
        640,
        666,
        692,
        718,
        744,
        770,
        796,
        822,
        848,
        874,
        900,
    ):
        constrain(constraints, f"PV_Base!D{_r}", unit_rate)

    for _r in (
        9,
        51,
        77,
        102,
        126,
        150,
        174,
        198,
        232,
        258,
        284,
        310,
        336,
        362,
        388,
        414,
        440,
        466,
        492,
        518,
        544,
        570,
        596,
        622,
        648,
        674,
        700,
        726,
        752,
        778,
        804,
        830,
        856,
        882,
    ):
        constrain(constraints, f"PV_Base!D{_r}", Literal[100])

    constrain(constraints, "PV_Base!D40", Literal[3])
    constrain(constraints, "PV_Base!D49", financial_type)
    for _dr in (
        69,
        71,
        72,
        120,
        122,
        144,
        146,
        168,
        192,
        194,
        216,
        218,
        250,
        252,
    ):
        constrain(constraints, f"PV_Base!D{_dr}", financial_type)
    # AM:BP bands mix blank leaves with formula rows (e.g. AM172); add domains per DynamicRefError.
    constrain(constraints, "PV_Base!BE158:BE176", financial_type)
    constrain(constraints, "PV_Base!AD188:BX188", financial_type)
    constrain(constraints, "PV_Base!BM212:CC212", financial_type)
    constrain(constraints, "PV_Base!AQ65:BD65", financial_type)
    constrain(constraints, "PV_Base!BC85:BP99", financial_type)
    for _r in (80, 88, 95, 97, 98, 99, 105):
        constrain(constraints, f"PV_Base!D{_r}", unit_rate)

    # B column = formulas from Input 4; constrain Input 4 G/H instead.


_constrain_pv_stress_and_pv_base_index_cells(LicDsfConstraints)


# ---------------------------------------------------------------------------
# PV_LC_NR1 / PV_LC_NR2 / PV_LC_NR3 constraints (local-currency new-loan output blocks)
# ---------------------------------------------------------------------------

def _constrain_pv_lc_nr(constraints: type[Any], sheet: str) -> None:
    financial_type = Annotated[float | None, RealBetween(0, 1e15)]

    # C28: text label "Stock of debt (in LC)"
    constrain(constraints, f"{sheet}!C28", Literal["Stock of debt (in LC)"])

    # Y7:BG7: interest rate (local currency) — unit rate; tail past F:U through AX/BD (OFFSET targets)
    _interest_lc_unit = Annotated[float | None, RealBetween(0, 1)]
    _lc7_min_col, _, _lc7_max_col, _ = range_boundaries("Y7:BG7")
    for _ci in range(_lc7_min_col, _lc7_max_col + 1):
        constrain(constraints, f"{sheet}!{get_column_letter(_ci)}7", _interest_lc_unit)

    # BC3:BD3 / BC5:BD5 are leaves only on PV_LC_NR1; NR2/NR3 use formulas in those cells.
    if sheet == "PV_LC_NR1":
        constrain(constraints, f"{sheet}!BC3:BD3", Literal[None])
        _bc5_col, _, _bd5_col, _ = range_boundaries("BC5:BD5")
        for _ci in range(_bc5_col, _bd5_col + 1):
            constrain(constraints, f"{sheet}!{get_column_letter(_ci)}5", financial_type)
    _y6_min_col, _, _y6_max_col, _ = range_boundaries("Y6:BD6")
    for _ci in range(_y6_min_col, _y6_max_col + 1):
        constrain(constraints, f"{sheet}!{get_column_letter(_ci)}6", financial_type)

    if sheet == "PV_LC_NR1":
        _r4_min_col, _, _r4_max_col, _ = range_boundaries("BC4:BD4")
        for _ci in range(_r4_min_col, _r4_max_col + 1):
            constrain(constraints, f"{sheet}!{get_column_letter(_ci)}4", financial_type)

    # AF column: "cumulative" rows in each 19-row loan output block
    for _r in range(26, 407, 19):
        constrain(constraints, f"{sheet}!AF{_r}", financial_type)

    # BB column: "Total debt service (in USD)" rows in each 19-row block
    for _r in range(30, 411, 19):
        constrain(constraints, f"{sheet}!BB{_r}", financial_type)

    # D column: three sub-types per 19-row block starting at row 23
    for _block_start in range(23, 404, 19):
        # offset 0: counter/start year (literal 0)
        constrain(constraints, f"{sheet}!D{_block_start}", Literal[0])
        # offset 5: stock row is a formula in the template (not a leaf).
        # offset 8: interest in USD (empty in D column)
        constrain(constraints, f"{sheet}!D{_block_start + 8}", financial_type)


_constrain_pv_lc_nr(LicDsfConstraints, "PV_LC_NR1")
_constrain_pv_lc_nr(LicDsfConstraints, "PV_LC_NR2")
_constrain_pv_lc_nr(LicDsfConstraints, "PV_LC_NR3")

# ---------------------------------------------------------------------------
# Input 1 - Basics
# ---------------------------------------------------------------------------

# enrichment_audit.json: first projection year; discount rate (template 0.05); ext/dom
# definition (data validation lookup!X4:X5).
constrain(LicDsfConstraints, "'Input 1 - Basics'!C18", Annotated[int, Between(1990, 2100)])
constrain(LicDsfConstraints, "'Input 1 - Basics'!C25", Annotated[float, RealBetween(0, 1)])
constrain(
    LicDsfConstraints,
    "'Input 1 - Basics'!C33",
    Literal["Residency-based", "Currency-based"],
)

# ---------------------------------------------------------------------------
# Input 3 - Macro-Debt data (DMX)
# ---------------------------------------------------------------------------

_INPUT3_DMX_A1_RANGES: tuple[str, ...] = (
    "AB100:AQ100",
    "AB101:AQ108",
    "AB109:AQ109",
    "AB111:AQ111",
    "AB112:AQ112",
    "AB113:AQ113",
    "AB114:AQ115",
    "AB116:AQ116",
    "AB117:AQ119",
    "AB120:AQ120",
    "AB122:AQ122",
    "AB123:AQ123",
    "AB124:AQ124",
    "AB125:AQ125",
    "AB126:AQ126",
    "AB128:AQ128",
    "AB129:AQ131",
    "AB12:AQ13",
    "AB132:AQ132",
    "AB133:AQ133",
    "AB134:AQ139",
    "AB141:AQ144",
    "AB147:AQ147",
    "AB149:AQ150",
    "AB152:AQ154",
    "AB155:AQ155",
    "AB156:AQ156",
    "AB157:AQ157",
    "AB166:AQ169",
    "AB172:AQ173",
    "AB175:AQ175",
    "AB176:AQ176",
    "AB177:AQ178",
    "AB179:AQ179",
    "AB180:AQ180",
    "AB19:AQ20",
    "AB22:AQ22",
    "AB24:AQ24",
    "AB26:AQ27",
    "AB28:AQ29",
    "AB30:AQ30",
    "AB34:AQ35",
    "AB38:AQ38",
    "AB41:AQ41",
    "AB43:AQ43",
    "AB52:AQ52",
    "AB55:AQ55",
    "AB57:AQ59",
    "AB65:AQ65",
    "AB66:AQ69",
    "AB70:AQ70",
    "AB71:AQ71",
    "AB72:AQ72",
    "AB73:AQ73",
    "AB74:AQ74",
    "AB75:AQ75",
    "AB76:AQ82",
    "AB83:AQ83",
    "AB84:AQ84",
    "AB85:AQ85",
    "AB86:AQ86",
    "AB87:AQ87",
    "AB88:AQ88",
    "AB89:AQ89",
    "AB90:AQ91",
    "AB94:AQ94",
    "AB98:AQ98",
    "AB92:AQ92",
    "AB93:AQ93",
    "AB95:AQ95",
    "AR100",
    "AR101:AR108",
    "AR109",
    "AR111",
    "AR112",
    "AR113",
    "AR114:AR115",
    "AR116",
    "AR117:AR119",
    "AR120",
    "AR122",
    "AR123",
    "AR124",
    "AR125",
    "AR126",
    "AR128",
    "AR129:AR131",
    "AR12:AR13",
    "AR132",
    "AR133",
    "AR134:AR139",
    "AR141:AR144",
    "AR147",
    "AR149:AR150",
    "AR152:AR154",
    "AR155",
    "AR156",
    "AR157",
    "AR166:AR169",
    "AR172:AR173",
    "AR175",
    "AR176",
    "AR177:AR178",
    "AR179",
    "AR180",
    "AR19:AR20",
    "AR22",
    "AR24",
    "AR26:AR27",
    "AR28:AR29",
    "AR30",
    "AR34:AR35",
    "AR38",
    "AR41",
    "AR43",
    "AR52",
    "AR55",
    "AR57:AR59",
    "AR65",
    "AR66:AR69",
    "AR70",
    "AR71",
    "AR73",
    "AR75",
    "AR72",
    "AR74",
    "AR76:AR82",
    "AR83",
    "AR84",
    "AR85",
    "AR86",
    "AR87",
    "AR88",
    "AR89",
    "AR90:AR91",
    "AR92",
    "AR93",
    "AR94",
    "AR95",
    "AR98",
    "BP65",
    "BP66:BP69",
    "BP70",
    "BP73",
    "BP75",
    "BP72",
    "BP74",
    "BP76:BP82",
    "BP83",
    "BP84",
    "BP85",
    "BP86",
    "BP87",
    "BP88",
    "BP89",
    "BP90",
    "BP91",
    "BP94",
    "BP92",
    "BP93",
    "M12:M13",
    "M35",
    "N12:N13",
    "O12:U13",
    "N142",
    "N166:N167",
    "N20",
    "N34:N35",
    "N41",
    "N43",
    "N53",
    "N59",
    "V12:V13",
    "V20",
    "V35",
    "W12:W13",
    "W138:W139",
    "W142",
    "W161:W164",
    "W166:W167",
    "W19:W20",
    "W34:W35",
    "W41",
    "W43",
    "W51:W53",
    "W55",
    "W57:W59",
    "X100",
    "X101:X108",
    "X109",
    "X111",
    "X112",
    "X113",
    "X114:X115",
    "X116",
    "X117:X119",
    "X120",
    "X122",
    "X123",
    "X124",
    "X125",
    "X126",
    "X128",
    "X129:X131",
    "X12:X13",
    "X132",
    "X133",
    "X134:X139",
    "X141:X144",
    "X147",
    "X149:X150",
    "X152:X154",
    "X154:X155",
    "X156",
    "X157",
    "X166:X169",
    "X172:X173",
    "X175",
    "X176",
    "X177:X178",
    "X179",
    "X180",
    "X19:X20",
    "X22",
    "X24",
    "X26:X27",
    "X28:X29",
    "X30",
    "X35",
    "X41",
    "X52",
    "X55",
    "X57:X58",
    "X65",
    "X66:X69",
    "X70",
    "X71",
    "X73",
    "X75",
    "X72",
    "X74",
    "X76:X82",
    "X83",
    "X84",
    "X85",
    "X86",
    "X87",
    "X88",
    "X89",
    "X90",
    "X91",
    "X94",
    "X98",
    "X92",
    "X93",
    "X95",
    "Y100:AA100",
    "Y101:AA108",
    "Y109:AA109",
    "Y111:AA111",
    "Y112:AA112",
    "Y113:AA113",
    "Y114:AA115",
    "Y116:AA116",
    "Y117:AA119",
    "Y120:AA120",
    "Y122:AA122",
    "Y123:AA123",
    "Y124:AA124",
    "Y125:AA125",
    "Y126:AA126",
    "Y128:AA132",
    "Y12:AA13",
    "Y133:AA133",
    "Y134:AA139",
    "Y141:AA144",
    "Y147:AA147",
    "Y149:AA150",
    "Y152:AA154",
    "Y155:AA155",
    "Y156:AA156",
    "Y157:AA157",
    "Y166:AA169",
    "Y172:AA173",
    "Y175:AA175",
    "Y176:AA176",
    "Y177:AA178",
    "Y179:AA179",
    "Y180:AA180",
    "Y19:AA20",
    "Y22:AA22",
    "Y24:AA24",
    "Y26:AA27",
    "Y28:AA29",
    "Y30:AA30",
    "Y34:AA35",
    "Y38:AA38",
    "Y41:AA41",
    "Y43:AA43",
    "Y52:AA52",
    "Y55:AA55",
    "Y57:AA59",
    "Y65:AA65",
    "Y66:AA69",
    "Y70:AA70",
    "Y71:AA71",
    "Y73:AA73",
    "Y75:AA75",
    "Y72:AA72",
    "Y74:AA74",
    "Y76:AA82",
    "Y83:AA83",
    "Y84:AA84",
    "Y85:AA85",
    "Y86:AA86",
    "Y87:AA87",
    "Y88:AA88",
    "Y89:AA89",
    "Y90:AA91",
    "Y94:AA94",
    "Y98:AA98",
    "Y89:AA89",
    "Y92:AA92",
    "Y93:AA93",
    "Y95:AA95",
)


def _constrain_input3_dmx(constraints: type[Any]) -> None:
    """Input 3 DMX macro series feeding INDEX (enrichment_audit: flows/GDP; may be negative)."""
    dmx_macro = Annotated[float | None, RealBetween(-1e15, 1e15)]
    q = "'Input 3 - Macro-Debt data(DMX)'"
    for a1 in _INPUT3_DMX_A1_RANGES:
        constrain(constraints, f"{q}!{a1}", dmx_macro)
    # Columns O:BZ: DMX grid outside AB:AQ audit ranges—see `_apply_lic_dsf_workbook_leaf_overlays`.


_constrain_input3_dmx(LicDsfConstraints)

# ---------------------------------------------------------------------------
# Input 4 - External Financing
# ---------------------------------------------------------------------------


def _constrain_input4_external_financing(constraints: type[Any]) -> None:
    """External financing (enrichment_audit: AG–AN and L–N flows; F interest; G grace; H maturity)."""
    financial_type = Annotated[float | None, RealBetween(0, 1e15)]
    unit_rate = Annotated[float | None, RealBetween(0, 1)]
    q = "'Input 4 - External Financing'"
    constrain(constraints, f"{q}!L10:N10", financial_type)
    # L–AT ladder mixes blanks with formula cells (e.g. R18); add ranges when DynamicRefError lists them.
    for a1 in (
        "AG10:AM10",
        "AG11:AM17",
        "AG19:AM19",
        "AG21:AM21",
        "AG22:AM22",
        "AG23:AM23",
        "AG26:AM26",
        "AG27:AM28",
        "AG29:AM29",
        "AG30:AM30",
        "AG32:AM32",
        "AG33:AM34",
        "AG35:AM35",
        "AG36:AM36",
        "AG38:AM42",
    ):
        constrain(constraints, f"{q}!{a1}", financial_type)
    # Columns AN:BP — projection years after AM (row 66 +1 ladder through BP; same 67/69–75 pattern per column).
    _an_col, _, _bp_col, _ = range_boundaries("AN1:BP1")
    for _ci in range(_an_col, _bp_col + 1):
        _pcol = get_column_letter(_ci)
        constrain(constraints, f"{q}!{_pcol}10:{_pcol}65", financial_type)
        constrain(constraints, f"{q}!{_pcol}67", unit_rate)
        constrain(constraints, f"{q}!{_pcol}69:{_pcol}70", financial_type)
        constrain(constraints, f"{q}!{_pcol}73:{_pcol}74", financial_type)
        constrain(constraints, f"{q}!{_pcol}75", unit_rate)

    for cell in ("F10", "F19", "F21", "F22", "F23", "F45"):
        constrain(constraints, f"{q}!{cell}", unit_rate)
    constrain(constraints, f"{q}!F18:F36", unit_rate)
    constrain(constraints, f"{q}!F38:F42", unit_rate)
    # "As of April 2018" creditor block: F holds maturity-style integers (not the same as F10 unit rates).
    constrain(constraints, f"{q}!F67:F75", Annotated[int | None, Between(1, 100)])
    constrain(
        constraints,
        f"{q}!E67:E75",
        Annotated[int | None, Between(0, 50)],
    )
    constrain(constraints, f"{q}!D10", financial_type)
    constrain(constraints, f"{q}!D12:D64", financial_type)
    constrain(constraints, f"{q}!D16", Literal["IDA NEW Blend floating"])
    constrain(constraints, f"{q}!D66:D73", financial_type)
    constrain(constraints, f"{q}!D75", financial_type)
    constrain(constraints, f"{q}!D76:D87", financial_type)
    constrain(constraints, f"{q}!C65:C73", financial_type)
    constrain(constraints, f"{q}!C74", Literal["IDA NEW Blend fixed"])
    constrain(constraints, f"{q}!C75", financial_type)
    constrain(constraints, f"{q}!C76:C90", financial_type)
    constrain(constraints, f"{q}!B11", Literal["IDA - regular"])
    constrain(constraints, f"{q}!B12", Literal["IDA - 50Y loans"])
    constrain(constraints, f"{q}!B13", Literal["IDA - SML"])
    constrain(constraints, f"{q}!B14", Literal["IDA NEW 40-year credits"])
    constrain(constraints, f"{q}!B15", Literal["IDA NEW Regular"])
    constrain(constraints, f"{q}!B16", Literal["IDA NEW Blend (also enter) -->"])
    constrain(constraints, f"{q}!B17", Literal["IDA NEW 60-year credits"])
    constrain(constraints, f"{q}!B67", Literal["IDA - small economy"])
    constrain(constraints, f"{q}!B68", Literal["IDA - regular"])
    constrain(constraints, f"{q}!B69", Literal["IDA - blend"])
    constrain(constraints, f"{q}!B70", Literal["IDA - SML"])
    constrain(constraints, f"{q}!B71", Literal["IDA - 50Y loans"])
    constrain(constraints, f"{q}!B72", Literal["IDA NEW 40-year credits"])
    constrain(constraints, f"{q}!B73", Literal["IDA NEW Regular"])
    constrain(constraints, f"{q}!B74", Literal["IDA NEW Blend (also enter) -->"])
    constrain(constraints, f"{q}!B75", Literal["IDA NEW 60-year credits"])
    constrain(constraints, f"{q}!B76:B95", financial_type)

    # G/H: PV_Base B2n/B2n+1 pairs reference the same Input 4 row for G/H; denominators use (H−G).
    # Literals match the template data_only snapshot; structurally blank grace/maturity rows use small
    # integers so (H−G) is never zero under numeric abstract analysis.
    for _gr, _gv, _hv in (
        (10, 5, 10),
        (11, 6, 38),
        (12, 10, 50),
        (13, 6, 12),
        (14, 11, 40),
        (15, 6, 31),
        (16, 5, 25),
        (17, 20, 60),
        (18, 5, 30),
        (19, 5, 30),
        (21, 5, 20),
        (22, 5, 25),
        (23, 5, 30),
        (26, 7, 20),
        (30, 5, 15),
        (32, 5, 15),
        (33, 5, 15),
        (34, 5, 15),
        (35, 5, 15),
        (36, 5, 15),
        (38, 9, 12),
        (39, 3, 12),
        (40, 1, 5),
        (41, 1, 5),
        (42, 1, 5),
        (54, 1, 2),
        (55, 3, 5),
        (56, 6, 7),
        (58, 0, 1),
        (59, 1, 2),
        (60, 3, 5),
        (61, 6, 7),
        # Template blanks; literals keep (H−G) strictly positive for PV_Base B-pair denominators.
        (27, 0, 2),
        (28, 0, 2),
        (29, 1, 2),
        (57, 0, 2),
    ):
        constrain(constraints, f"{q}!G{_gr}", LiteralType[_gv])
        constrain(constraints, f"{q}!H{_gr}", LiteralType[_hv])


_constrain_input4_external_financing(LicDsfConstraints)

# ---------------------------------------------------------------------------
# Input 5 - Local-debt Financing
# ---------------------------------------------------------------------------


def _constrain_input5_local_debt(constraints: type[Any]) -> None:
    """Domestic debt instruments: grace/maturity (C/D), interest by year (I–AA on assumption rows),
    issuance and adjustment flows (enrichment_audit + template row 5–7 headers)."""

    q = "'Input 5 - Local-debt Financing'"
    financial = Annotated[float | None, RealBetween(0, 1e15)]
    financial_signed = Annotated[float | None, RealBetween(-1e15, 1e15)]
    unit_rate = Annotated[float | None, RealBetween(0, 1)]
    grace = Annotated[int | None, Between(0, 50)]
    maturity = Annotated[int | None, Between(1, 100)]

    constrain(constraints, f"{q}!C16:C22", grace)
    constrain(constraints, f"{q}!C10", grace)
    # Instrument block C values: template is mostly 0; literals keep OFFSET fallback enumeration small.
    for row in (83, 86, 93, 94, 95, 100, 101, 104, 105, 106, 108, 109, 110):
        constrain(constraints, f"{q}!C{row}", Literal[0])
    constrain(constraints, f"{q}!C89", Literal[0.2])
    constrain(constraints, f"{q}!C90", Literal[0.1])
    constrain(constraints, f"{q}!C91", Literal[0.05])

    constrain(constraints, f"{q}!C78", Annotated[int | None, Between(0, 1)])

    constrain(constraints, f"{q}!D16:D22", maturity)
    # D10 and the instrument block: template uses 1, 0, small decimals, or blank—same enumeration
    # issue as column C if each cell keeps Between(1, 100) maturity (100^N branches).
    constrain(constraints, f"{q}!D10", Literal[1])
    for row in (83, 86, 93, 94, 95, 100, 101, 104, 105, 106, 108, 109, 110):
        constrain(constraints, f"{q}!D{row}", Literal[0])
    constrain(constraints, f"{q}!D89", Literal[0.2])
    constrain(constraints, f"{q}!D90", Literal[0.1])
    constrain(constraints, f"{q}!D91", Literal[0.05])
    constrain(constraints, f"{q}!D92", Literal[None])

    # Main instrument ladder E/F: template is all zeros; literals avoid 11^N small-int enumeration.
    for row in (93, 94, 95, 100, 101, 104, 105, 106, 108, 109, 110):
        constrain(constraints, f"{q}!E{row}", Literal[0])
        constrain(constraints, f"{q}!F{row}", Literal[0])
    constrain(constraints, f"{q}!E84", Literal[None])
    constrain(constraints, f"{q}!F84", Literal[None])
    constrain(constraints, f"{q}!F87", Literal[None])
    constrain(constraints, f"{q}!F88", Literal[None])
    # Rows 83–92 instrument band: E/F are template-filled shares or zeros; literals match the empty
    # template and keep OFFSET subgraphs from multiplying wide numeric domains.
    constrain(constraints, f"{q}!E83", Literal[0])
    constrain(constraints, f"{q}!F83", Literal[0])
    constrain(constraints, f"{q}!E86", Literal[0])
    constrain(constraints, f"{q}!F86", Literal[0])
    constrain(constraints, f"{q}!E89", Literal[0.19])
    constrain(constraints, f"{q}!F89", Literal[0.18])
    constrain(constraints, f"{q}!E90", Literal[0.15])
    constrain(constraints, f"{q}!F90", Literal[0.2])
    constrain(constraints, f"{q}!E91", Literal[0.1])
    constrain(constraints, f"{q}!F91", Literal[0.2])
    constrain(constraints, f"{q}!E92", Literal[None])
    constrain(constraints, f"{q}!F92", Literal[None])

    constrain(constraints, f"{q}!I16:N22", unit_rate)
    constrain(constraints, f"{q}!J10", unit_rate)
    constrain(constraints, f"{q}!K10", unit_rate)

    for col_idx in range(9, 30):  # I:AC — adjustment row (signed flows including SoE removal)
        constrain(constraints, f"{q}!{get_column_letter(col_idx)}63", financial_signed)

    for addr in (
        "AD93",
        "AD94",
        "AD95",
        "AD108",
        "AD109",
        "AD110",
        "AD188",
        "AD191",
        "AD193",
    ):
        constrain(constraints, f"{q}!{addr}", financial)

    for row in (93, 94, 95, 100, 101, 104, 105, 106, 108, 109, 110, 250, 254, 274, 278, 298, 302, 322, 392, 461):
        constrain(constraints, f"{q}!AE{row}", financial)

    for row in (93, 94, 95, 100, 101, 104, 105, 106, 108, 109, 110, 250, 274, 298, 322, 392, 461, 488, 512):
        constrain(constraints, f"{q}!AF{row}", financial)

    # AG:BT projection grids are mostly formulas in the template; constrain true OFFSET leaves via DynamicRefError.

    # Column M: blank INPUT cells between formula ladders on rows 465–651 (OFFSET leaves).
    for _m_lo, _m_hi in (
        (465, 467),
        (473, 491),
        (497, 516),
        (522, 538),
        (544, 560),
        (566, 586),
        (592, 608),
        (614, 630),
        (636, 651),
    ):
        constrain(constraints, f"{q}!M{_m_lo}:M{_m_hi}", financial)

    for _af_lo, _af_hi in (
        (227, 229),
        (232, 253),
        (255, 277),
        (279, 301),
        (303, 326),
        (329, 348),
        (351, 370),
        (373, 396),
        (399, 418),
        (421, 440),
        (443, 463),
        (465, 467),
        (469, 491),
        (493, 516),
        (519, 538),
        (541, 560),
        (563, 586),
        (589, 608),
        (611, 630),
        (633, 651),
    ):
        constrain(constraints, f"{q}!AF{_af_lo}:AF{_af_hi}", financial)
        # AE: projection column left of AF; same row bands (e.g. AE272 OFFSET leaves).
        constrain(constraints, f"{q}!AE{_af_lo}:AE{_af_hi}", financial)

    # AE blank cells in single-row gaps between AF ladder bands (AF is formula; AE not in AE:AF range).
    for _ae_gap in (231, 328, 350, 372, 398, 420, 442, 518, 540, 562, 588, 610, 632):
        constrain(constraints, f"{q}!AE{_ae_gap}", financial)

    # AY:BT bands mix formulas with blanks; BU blanks are filled in `_apply_lic_dsf_workbook_leaf_overlays`.

    for row in (230, 254, 278, 302, 327, 397):
        constrain(constraints, f"{q}!H{row}", financial)

    for row in (488, 581):
        for col_idx in range(9, 28):  # I:AA — issuance / projection inputs (leaf rows only here)
            constrain(constraints, f"{q}!{get_column_letter(col_idx)}{row}", financial)

    for row in (250, 274, 298, 322, 439, 440, 488, 512, 581):
        constrain(constraints, f"{q}!AB{row}", financial)

    # BU: hundreds of structural blanks between formula rows—see `_apply_lic_dsf_workbook_leaf_overlays`.


_constrain_input5_local_debt(LicDsfConstraints)

# ---------------------------------------------------------------------------
# Input 6 (Tailored / optional Standard Test) and Input 8 (SDR)
# ---------------------------------------------------------------------------

def _constrain_input6_input8(constraints: type[Any]) -> None:
    """Tailored and standardized stress options; SDR sheet (enrichment_audit + template dropdowns)."""
    _threshold = Literal["Historical average only", "Baseline projection only", "Whichever is lower"]
    financial = Annotated[float | None, RealBetween(0, 1e15)]
    financial_signed = Annotated[float | None, RealBetween(-1e15, 1e15)]
    unit_rate = Annotated[float | None, RealBetween(0, 1)]

    q6t = "'Input 6 - Tailored Tests'"
    q6o = "'Input 6(optional)-Standard Test'"
    q8 = "'Input 8 - SDR'"

    constrain(constraints, f"{q6t}!C6", Literal["On", "Off"])

    constrain(constraints, f"{q6o}!C4", Literal["New", "Old"])
    constrain(constraints, f"{q6o}!C5", _threshold)
    constrain(constraints, f"{q6o}!C7", _threshold)
    constrain(constraints, f"{q6o}!C8", Literal["On", "Off"])
    constrain(constraints, f"{q6o}!C17", Annotated[float, RealBetween(0, 10)])
    constrain(constraints, f"{q6o}!D8", Literal[None])
    constrain(constraints, f"{q6o}!D9", Literal[None])
    constrain(constraints, f"{q6o}!D18", _threshold)

    constrain(constraints, f"{q8}!B6:B7", financial)
    constrain(constraints, f"{q8}!C11:C12", financial_signed)
    constrain(constraints, f"{q8}!D11:V12", financial_signed)
    constrain(constraints, f"{q8}!D14:V14", financial_signed)
    constrain(constraints, f"{q8}!W14", unit_rate)
    constrain(constraints, f"{q8}!AG37", financial)
    constrain(constraints, f"{q8}!J27", Literal[None])
    constrain(constraints, f"{q8}!S37", Literal[None])
    constrain(constraints, f"{q8}!X27", financial)
    constrain(constraints, f"{q8}!Y28", financial)

    # Blend fixed/floating scenario labels (referenced by OFFSET/INDIRECT argument subgraph).
    q_blend = "'BLEND floating calculations WB'"
    _blend_fin = Annotated[float | None, RealBetween(0, 1e15)]
    _blend_unit = Annotated[float | None, RealBetween(0, 1)]
    constrain(constraints, f"{q_blend}!B5", Literal["IDA NEW Blend fixed"])
    constrain(constraints, f"{q_blend}!C5", Literal[None])
    constrain(constraints, f"{q_blend}!D5", _blend_unit)
    constrain(constraints, f"{q_blend}!E5", _blend_fin)
    constrain(constraints, f"{q_blend}!F5", _blend_fin)
    constrain(constraints, f"{q_blend}!B6", Literal["IDA NEW Blend floating"])
    constrain(constraints, f"{q_blend}!C6", Literal["USD"])
    constrain(constraints, f"{q_blend}!E6", _blend_fin)
    constrain(constraints, f"{q_blend}!F6", _blend_fin)
    constrain(constraints, f"{q_blend}!L6", _blend_unit)
    constrain(constraints, f"{q_blend}!M6", Literal[None])
    constrain(constraints, f"{q_blend}!M7", Literal[None])
    constrain(
        constraints,
        f"{q_blend}!M8",
        Literal["Linear interpolation swap curve"],
    )
    constrain(constraints, f"{q_blend}!M9", Literal["Year"])
    for _blend_m_r, _blend_m_v in zip(range(10, 15), range(1, 6), strict=True):
        constrain(constraints, f"{q_blend}!M{_blend_m_r}", Literal[_blend_m_v])  # ty: ignore[invalid-type-form]
    for _blend_m_r, _blend_m_v in zip(range(15, 40), range(6, 31), strict=True):
        constrain(constraints, f"{q_blend}!M{_blend_m_r}", Literal[_blend_m_v])  # ty: ignore[invalid-type-form]
    constrain(constraints, f"{q_blend}!O6", Literal[None])
    constrain(constraints, f"{q_blend}!O7", Literal[None])
    constrain(constraints, f"{q_blend}!O8", Literal[None])
    constrain(constraints, f"{q_blend}!O9", Literal["Linear interpolation"])
    # O10:O39: interpolated swap rates (array formulas); domains are applied in
    # `_apply_lic_dsf_workbook_leaf_overlays` for dynamic ref resolution.


_constrain_input6_input8(LicDsfConstraints)

# ---------------------------------------------------------------------------
# Translation table constraints
# ---------------------------------------------------------------------------

# Translation labels referenced by dynamic formulas (OFFSET/INDIRECT).
# Column C = English, D–F = other languages (Spanish, Portuguese, French per workbook layout).
# ---------------------------------------------------------------------------
# Ext_Debt_Data constraints
# ---------------------------------------------------------------------------

# AA403:AG403 — "Exchange rate (pa)" projection columns (years); may also map to
# creditor-row financial data depending on workbook layout.
constrain(LicDsfConstraints, "Ext_Debt_Data!AA403:AG403", Annotated[float | None, RealBetween(0, 1e15)])

# F383:F384 — short-term debt principal / interest (or exchange rate in some layouts)
constrain(LicDsfConstraints, "Ext_Debt_Data!F383:F384", Annotated[float | None, RealBetween(0, 1e15)])

constrain(LicDsfConstraints, "Ext_Debt_Data!BO79:CF79", Annotated[float | None, RealBetween(0, 1e15)])

# ---------------------------------------------------------------------------
# Translation table constraints
# ---------------------------------------------------------------------------

constrain(LicDsfConstraints, "translation!C90", Literal["Residency-based"])
constrain(LicDsfConstraints, "translation!C451", Literal["Grace period"])
constrain(LicDsfConstraints, "translation!C452", Literal["Loan Maturity"])
constrain(LicDsfConstraints, "translation!C898", Literal["Projections"])

constrain(LicDsfConstraints, "translation!D451", Literal["Período de gracia"])
constrain(LicDsfConstraints, "translation!E451", Literal["Prazo de carência"])
constrain(LicDsfConstraints, "translation!F451", Literal["Différé d'amortissement "])

constrain(LicDsfConstraints, "translation!D452", Literal["Vencimiento del préstamo"])
constrain(LicDsfConstraints, "translation!E452", Literal["Vencimento do empr."])
constrain(LicDsfConstraints, "translation!F452", Literal["Échéance  crédit "])

constrain(LicDsfConstraints, "translation!D898", Literal["Projections"])
constrain(LicDsfConstraints, "translation!E898", Literal["Projecções"])
constrain(LicDsfConstraints, "translation!F898", Literal["Proyecciones"])

def _workbook_cell_raw_is_formula(raw: object) -> bool:
    """Match :func:`create_dependency_graph` formula detection (string or ArrayFormula)."""
    if isinstance(raw, ArrayFormula):
        text = raw.text or ""
        if text and not text.startswith("="):
            text = f"={text}"
        return isinstance(text, str) and text.startswith("=")
    return isinstance(raw, str) and raw.startswith("=")


def _apply_lic_dsf_workbook_leaf_overlays(constraints: type[Any]) -> None:
    """Add domains for OFFSET/INDIRECT leaves only (skip template formula cells).

    Ranges mirror where the workbook leaves empty or numeric inputs next to dynamic formulas.
    Each `add_range` / `add_cell` pass skips formulas so constraints stay on true inputs only.
    """
    if not WORKBOOK_PATH.is_file():
        return

    # Generic non-negative monetary / stock-flow magnitudes (template scale); None allows blanks.
    financial = Annotated[float | None, RealBetween(0, 1e15)]
    keep_vba = WORKBOOK_PATH.suffix.lower() == ".xlsm"
    wb = fastpyxl.load_workbook(WORKBOOK_PATH, data_only=False, keep_vba=keep_vba)
    try:

        def add_range(sheet: str, range_a1: str, ann: Any) -> None:
            if sheet not in wb.sheetnames:
                return
            ws = wb[sheet]
            for key in cells_in_range(sheet, range_a1):
                _, coord = parse_range_spec(key)
                if not _workbook_cell_raw_is_formula(ws[coord].value):
                    constrain(constraints, key, ann)

        def add_cell(sheet: str, coord: str, ann: Any) -> None:
            if sheet not in wb.sheetnames:
                return
            raw = wb[sheet][coord].value
            if _workbook_cell_raw_is_formula(raw):
                return
            constrain(constraints, format_key(sheet, coord), ann)

        def add_range_with_formula_alias(
            sheet: str, formula_alias: str, range_a1: str, ann: Any
        ) -> None:
            if sheet not in wb.sheetnames:
                return
            ws = wb[sheet]
            for key in cells_in_range(sheet, range_a1):
                _, coord = parse_range_spec(key)
                if _workbook_cell_raw_is_formula(ws[coord].value):
                    continue
                constrain(constraints, key, ann)
                constrain(constraints, format_key(formula_alias, coord), ann)

        def add_cell_with_formula_alias(
            sheet: str, formula_alias: str, coord: str, ann: Any
        ) -> None:
            if sheet not in wb.sheetnames:
                return
            raw = wb[sheet][coord].value
            if _workbook_cell_raw_is_formula(raw):
                return
            constrain(constraints, format_key(sheet, coord), ann)
            constrain(constraints, format_key(formula_alias, coord), ann)

        # C4 sheet: stress-test layout. Most C–G cells are blank or fixed labels; literals match
        # template text so INDIRECT/OFFSET resolution does not treat them as unconstrained strings.
        _c4 = "C4_Market_financing"
        _c4_formula_alias = "Market_financing"
        add_range_with_formula_alias(_c4, _c4_formula_alias, "C4:C53", Literal[None])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "D4:D77", Literal[None])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "E4:G53", Literal[None])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "D20:F20", Literal["Historical "])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "D21:F21", Literal["Average "])
        add_range_with_formula_alias(
            _c4,
            _c4_formula_alias,
            "E33",
            Literal["Maturity - Grace (to determine bullet / amortization)"],
        )
        add_range_with_formula_alias(_c4, _c4_formula_alias, "E34", Literal["Bullet (1) or Amort. (>1)"])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "F33", Literal["Stress test"])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "F34", Literal["Maturity"])
        add_range_with_formula_alias(_c4, _c4_formula_alias, "G34", Literal["Grace"])
        # AB column: C4 scenario controls (binary switches, blanks, and numeric shocks). Domains are
        # wide enough for typical stress magnitudes; template defaults include e.g. 15, 0.3, 400, 5.
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB20", Literal[0, 1])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB21", Literal[None])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB22", Annotated[float, RealBetween(0, 100)])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB23", Annotated[float, RealBetween(0, 1)])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB24", Annotated[float, RealBetween(0, 2000)])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB25", Annotated[int, Between(1, 50)])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB26", Annotated[float, RealBetween(0, 1)])
        add_cell_with_formula_alias(_c4, _c4_formula_alias, "AB27", Annotated[float, RealBetween(0, 1)])
        # E51:G53 are formulas on the physical sheet (Ext_Debt_Data G–I columns) but dynamic-ref
        # expansion still requests domains for the legacy codename `Market_financing!…`.
        for _c4_efg_col in ("E", "F", "G"):
            for _c4_g_row in range(48, 54):
                constrain(
                    constraints,
                    format_key(_c4_formula_alias, f"{_c4_efg_col}{_c4_g_row}"),
                    financial,
                )

        # BLEND floating sheet: column O is the swap curve used in blend calculations; stored as array
        # formulas but the grapher still needs a rate domain (decimal 0–1) per tenor row.
        _blend = "BLEND floating calculations WB"
        _swap_rate = Annotated[float | None, RealBetween(0, 1)]
        if _blend in wb.sheetnames:
            for _br in range(10, 40):
                constrain(constraints, format_key(_blend, f"O{_br}"), _swap_rate)

        # Input 8 - SDR: columns T–Z include many OFFSET-adjacent blanks outside D:V bands.
        _q8_ov = "Input 8 - SDR"
        if _q8_ov in wb.sheetnames:
            ws8 = wb[_q8_ov]
            _v_i8 = column_index_from_string("T")
            _z_i8 = column_index_from_string("Z")
            for _r8 in range(11, 45):
                for _ci8 in range(_v_i8, _z_i8 + 1):
                    _a8 = f"{get_column_letter(_ci8)}{_r8}"
                    _rv8 = ws8[_a8].value
                    if not _workbook_cell_raw_is_formula(_rv8):
                        _ann8: Any = Literal[None] if _rv8 is None else financial
                        constrain(constraints, format_key(_q8_ov, _a8), _ann8)

        # Input 3 DMX: columns O–BZ include macro inputs not listed in enrichment_audit AB:AQ ranges.
        _q3 = "Input 3 - Macro-Debt data(DMX)"
        _dmx_wide = Annotated[float | None, RealBetween(-1e15, 1e15)]
        if _q3 in wb.sheetnames:
            ws3 = wb[_q3]
            _o_i = column_index_from_string("O")
            _bz_i = column_index_from_string("BZ")
            for _r3 in range(1, 300):
                for _ci3 in range(_o_i, _bz_i + 1):
                    _dmx_a1 = f"{get_column_letter(_ci3)}{_r3}"
                    _rv3 = ws3[_dmx_a1].value
                    if not _workbook_cell_raw_is_formula(_rv3):
                        _ann3: Any = Literal[None] if _rv3 is None else _dmx_wide
                        constrain(constraints, format_key(_q3, _dmx_a1), _ann3)

        # Input 4: L–AT ladder and related cells are user entry points for external financing
        # projections; constrain as non-negative flows where the cell is a leaf.
        _q4 = "Input 4 - External Financing"
        for a1 in (
            "L18:Q18",
            "L20:Q20",
            "R18:W20",
            "X18:AT20",
            "L24:Q25",
            "R24:W25",
            "X24:AT24",
            "X25:AT25",
            "L31:Q31",
            "R31:W31",
            "X31:AT31",
            "AC31:AM31",
            "L37:Q37",
            "R37:W37",
            "X37:AT37",
            "L43:Q43",
            "M44:Q45",
            "L46:Q47",
        ):
            add_range(_q4, a1, financial)
        for addr in ("M11", "N14:N15", "M16:O16", "M17:O17"):
            add_range(_q4, addr, financial)
        # Creditor ladder column I: mostly structural blanks around row 66 formulas.
        add_range(_q4, "I62:I94", Literal[None])
        # Column O: same block mixes blanks, formulas, and template rate/scalar literals.
        add_range(_q4, "O62:O94", financial)
        # Columns N–AT: creditor / rate ladder leaves (per-cell skip for formulas).
        ws_i4 = wb[_q4]
        _n_i4 = column_index_from_string("N")
        _at_i4 = column_index_from_string("AT")
        for _r4x in range(62, 95):
            for _ci4x in range(_n_i4, _at_i4 + 1):
                _a4x = f"{get_column_letter(_ci4x)}{_r4x}"
                _rv4x = ws_i4[_a4x].value
                if not _workbook_cell_raw_is_formula(_rv4x):
                    _ann4x: Any = Literal[None] if _rv4x is None else financial
                    constrain(constraints, format_key(_q4, _a4x), _ann4x)

        # Input 5: wide domestic-debt projection grids (AG:BT and single-column ladders). Cells are
        # issuance, stock, or year-by-year debt-service inputs depending on row block headers.
        _q5 = "Input 5 - Local-debt Financing"
        for row in (93, 94, 95, 100, 101, 104, 105, 106, 108, 109, 110):
            add_range(_q5, f"AG{row}:AJ{row}", financial)
        for _lo, _hi in (
            (5, 20),
            (77, 78),
            (128, 132),
            (139, 193),
            (199, 206),
            (217, 220),
            (222, 224),
            (227, 463),
            (465, 651),
        ):
            add_range(_q5, f"AG{_lo}:AJ{_hi}", financial)
            add_range(_q5, f"AK{_lo}:BT{_hi}", financial)
            for _mid_col in "HIJKLMNOPQRSTUVWXYZ":
                add_range(_q5, f"{_mid_col}{_lo}:{_mid_col}{_hi}", financial)
        # AY/BA: anchor columns in some instrument blocks (totals or carry-downs).
        for row in (254, 278, 302, 392, 468, 492):
            add_cell(_q5, f"AY{row}", financial)
        for row in (250, 274, 298, 322, 392, 463):
            add_cell(_q5, f"BA{row}", financial)
        # BB:BT bands: projection rectangles aligned to the sheet’s domestic-instrument sections.
        _bb_bt_row_ranges: tuple[tuple[int, int], ...] = (
            (248, 253),
            (260, 277),
            (284, 301),
            (308, 326),
            (333, 348),
            (355, 370),
            (377, 396),
            (403, 418),
            (425, 467),
            (474, 491),
            (498, 516),
        )
        for lo, hi in _bb_bt_row_ranges:
            add_range(_q5, f"BB{lo}:BT{hi}", financial)
        # BU: beside the BB:BT ladder; template uses formulas on most rows and leaves the rest blank
        # for OFFSET targets—constrain every non-formula BU cell as empty (None).
        ws_i5 = wb[_q5]
        for _bur in range(1, 700):
            _bu_a1 = f"BU{_bur}"
            if not _workbook_cell_raw_is_formula(ws_i5[_bu_a1].value):
                constrain(constraints, format_key(_q5, _bu_a1), Literal[None])
        # AA:AZ (rows 220–659): ladder blanks to the left of AK:BT grids; OFFSET still references them.
        _aa_i = column_index_from_string("AA")
        _az_i = column_index_from_string("AZ")
        for _r in range(220, 660):
            for _ci in range(_aa_i, _az_i + 1):
                _cl = get_column_letter(_ci)
                _aa_az = f"{_cl}{_r}"
                _raw = ws_i5[_aa_az].value
                if not _workbook_cell_raw_is_formula(_raw):
                    _ann_aa: Any = Literal[None] if _raw is None else financial
                    constrain(constraints, format_key(_q5, _aa_az), _ann_aa)
        # Lower block: residual financing ladder (I461:AA520 mixes blanks and inputs).
        add_cell(_q5, "I461", financial)
        add_range(_q5, "I461:AA520", financial)

        # PV_stress_com: commercial debt under stress scenarios—D column mixes calendar years,
        # index bases (100), and cash-flow scalars; H:AE time-series; AR:BP debt-service grid
        # (per-cell skip keeps formula bands); AF/BD block totals.
        if "PV_stress_com" in wb.sheetnames:
            ws_ps = wb["PV_stress_com"]
            for r in range(9, 141):
                addr = f"D{r}"
                raw = ws_ps[addr].value
                if _workbook_cell_raw_is_formula(raw):
                    continue
                if r in (10, 22, 35):
                    ann: Any = Literal[2024]
                elif r in (23, 24, 28):
                    ann = Literal[100]
                else:
                    ann = financial
                constrain(constraints, format_key("PV_stress_com", addr), ann)
            _cols_ps = (
                "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
                "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE",
            )
            for col in _cols_ps:
                for r in range(36, 141):
                    addr = f"{col}{r}"
                    if not _workbook_cell_raw_is_formula(ws_ps[addr].value):
                        constrain(constraints, format_key("PV_stress_com", addr), financial)
            for r in range(37, 142):
                addr = f"AF{r}"
                if not _workbook_cell_raw_is_formula(ws_ps[addr].value):
                    constrain(constraints, format_key("PV_stress_com", addr), financial)
            for r in range(27, 132):
                addr = f"BD{r}"
                if not _workbook_cell_raw_is_formula(ws_ps[addr].value):
                    constrain(constraints, format_key("PV_stress_com", addr), financial)
            _ar_i = column_index_from_string("AR")
            _bp_ps = column_index_from_string("BP")
            for _r in range(27, 142):
                for _ci in range(_ar_i, _bp_ps + 1):
                    _addr_ar = f"{get_column_letter(_ci)}{_r}"
                    if not _workbook_cell_raw_is_formula(ws_ps[_addr_ar].value):
                        constrain(constraints, format_key("PV_stress_com", _addr_ar), financial)
            # Some OFFSET paths reference columns past AE; sweep AF:BG for remaining leaves.
            add_range("PV_stress_com", "AF36:BG140", financial)

        # PV_Base: AM:BP amortization / projection band; sweep leaves through row 900 (OFFSET targets).
        # Rows below 130 include structural blanks (e.g. BC75) beside formula ladders.
        if "PV_Base" in wb.sheetnames:
            ws_pb = wb["PV_Base"]
            _am_pb = column_index_from_string("AM")
            _bp_pb = column_index_from_string("BP")
            for _rpb in range(65, 901):
                for _cipb in range(_am_pb, _bp_pb + 1):
                    _apb = f"{get_column_letter(_cipb)}{_rpb}"
                    _rpbv = ws_pb[_apb].value
                    if not _workbook_cell_raw_is_formula(_rpbv):
                        _ann_pb: Any = Literal[None] if _rpbv is None else financial
                        constrain(constraints, format_key("PV_Base", _apb), _ann_pb)

        # Local-currency new loan sheets: D (block +5) stock input; AF:BG for OFFSET blanks (excludes
        # Y:AE where year-index ladders compare to column C—broad domains there explode IF fallbacks).
        _af_lc = column_index_from_string("AF")
        _bg_lc = column_index_from_string("BG")
        for _sheet in ("PV_LC_NR1", "PV_LC_NR2", "PV_LC_NR3"):
            if _sheet not in wb.sheetnames:
                continue
            ws_lc = wb[_sheet]
            for _block_start in range(23, 404, 19):
                addr = f"D{_block_start + 5}"
                if not _workbook_cell_raw_is_formula(ws_lc[addr].value):
                    constrain(constraints, format_key(_sheet, addr), financial)
            for _r in range(1, 411):
                for _ci in range(_af_lc, _bg_lc + 1):
                    _afb = f"{get_column_letter(_ci)}{_r}"
                    _rvb = ws_lc[_afb].value
                    if not _workbook_cell_raw_is_formula(_rvb):
                        _ann_lc: Any = Literal[None] if _rvb is None else financial
                        constrain(constraints, format_key(_sheet, _afb), _ann_lc)
    finally:
        wb.close()


_apply_lic_dsf_workbook_leaf_overlays(LicDsfConstraints)

# Those cells are array formulas; we still attach swap-rate domains for dynamic ref resolution.
_BLEND_O_CONSTRAINT_KEY = re.compile(
    r"^'BLEND floating calculations WB'!O(10|[1-3][0-9])$"
)
_MARKET_FINANCING_INDIRECT_MIRROR_KEY = re.compile(
    r"^Market_financing![EFG](4[89]|5[0-3])$"
)


def verify_lic_dsf_constraints_target_leaves(
    workbook_path: Path,
    constraints_type: type[Any],
) -> None:
    """Fail fast if any constrained address is a formula cell in the template.

    OFFSET/INDIRECT/INDEX resolution expects constraints on leaf inputs only.
    """
    if not workbook_path.is_file():
        return

    hints = get_type_hints(constraints_type, include_extras=True)
    if not hints:
        return

    keep_vba = workbook_path.suffix.lower() == ".xlsm"
    wb = fastpyxl.load_workbook(workbook_path, data_only=False, keep_vba=keep_vba)
    try:
        formula_cells: list[str] = []
        missing: list[str] = []
        for spec_key in hints:
            sheet_name, range_a1 = parse_range_spec(spec_key)
            for cell_key in cells_in_range(sheet_name, range_a1):
                sh, coord = parse_range_spec(cell_key)
                sh = _CONSTRAINT_VERIFY_SHEET_ALIASES.get(sh, sh)
                if sh not in wb.sheetnames:
                    missing.append(cell_key)
                    continue
                raw = wb[sh][coord].value
                if _workbook_cell_raw_is_formula(raw):
                    # Exception: swap curve O10:O39 (see overlay)—constrained despite array formula.
                    if _BLEND_O_CONSTRAINT_KEY.match(cell_key):
                        continue
                    # Legacy codename mirrors C4 formulas (see overlay); valid for dynamic-ref domain only.
                    if _MARKET_FINANCING_INDIRECT_MIRROR_KEY.match(cell_key):
                        continue
                    formula_cells.append(cell_key)
    finally:
        wb.close()

    if missing or formula_cells:
        parts: list[str] = []
        if formula_cells:
            sample = ", ".join(sorted(formula_cells)[:20])
            more = f" (+{len(formula_cells) - 20} more)" if len(formula_cells) > 20 else ""
            parts.append(
                "constrained cells that contain formulas (constraints must target leaves only): "
                f"{sample}{more}"
            )
        if missing:
            ms = ", ".join(sorted(missing)[:20])
            mmore = f" (+{len(missing) - 20} more)" if len(missing) > 20 else ""
            parts.append(f"constrained cells on missing sheets: {ms}{mmore}")
        raise ValueError(
            "LicDsfConstraints validation failed: " + "; ".join(parts) + "."
        )


def _get_missing_constraints(specs: list[str], constraints: type[Any]) -> list[str]:
    def _normalize_sheet(sheet: str) -> str:
        """Strip surrounding single-quotes so format_key can re-add them consistently."""
        if sheet.startswith("'") and sheet.endswith("'"):
            return sheet[1:-1]
        return sheet

    def expand_spec(spec: str) -> list[str]:
        if "!" not in spec:
            return [spec]
        sheet, range_part = spec.split("!", 1)
        sheet = _normalize_sheet(sheet)
        if ":" not in range_part:
            return [format_key(sheet, range_part)]

        # Handle ranges like Sheet!A1:Sheet!B2
        if ":" in range_part and "!" in range_part.split(":")[1]:
            parts = range_part.split(":")
            start_a1 = parts[0]
            end_a1 = parts[1].split("!", 1)[1]
        else:
            start_a1, end_a1 = range_part.split(":", 1)

        min_col, min_row, max_col, max_row = range_boundaries(f"{start_a1}:{end_a1}")
        cells = []
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cells.append(format_key(sheet, f"{get_column_letter(c)}{r}"))
        return cells

    missing = []
    for spec in specs:
        expanded = expand_spec(spec)
        for cell in expanded:
            if cell not in constraints.__annotations__:
                missing.append(spec)
                break
    return missing


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
    "'BLEND floating calculations WB'!B5",
    "'BLEND floating calculations WB'!B6",
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
