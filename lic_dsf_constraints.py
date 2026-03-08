#!/usr/bin/env python3
"""
Constraint-based config for LIC-DSF dynamic references (OFFSET/INDIRECT).

Used when building the dependency graph from lic-dsf-template-2026-01-31.xlsm.
Cells that feed OFFSET/INDIRECT are declared here with address-style keys; add
entries when the graph builder raises DynamicRefError.
"""

from __future__ import annotations

from typing import Literal, TypedDict

from excel_grapher.grapher import DynamicRefConfig

# Constraint types for cells that feed OFFSET/INDIRECT. Keys are address-style
# (e.g. "Sheet1!B1"). Add entries when the graph builder raises DynamicRefError.
class LicDsfConstraints(TypedDict, total=False):
    pass


# PV_Base!B9xx = CONCAT("$", A9xx, "$", $A$<row>) → INDIRECT($B9xx). Row-index cells A917, A941, A965 (fixed).
LicDsfConstraints.__annotations__["PV_Base!A917"] = Literal[64]
LicDsfConstraints.__annotations__["PV_Base!A941"] = Literal[90]
LicDsfConstraints.__annotations__["PV_Base!A965"] = Literal[115]
# A918:A938, A942:A962, A966:A986 each has a single cached letter D, E, …, X.
for _start, _end in [(918, 939), (942, 963), (966, 987)]:
    for _row in range(_start, _end):
        LicDsfConstraints.__annotations__[f"PV_Base!A{_row}"] = str
# B9xx holds the resulting ref string (e.g. "$D$64") consumed by INDIRECT.
for _start, _end, _anchor in [(918, 939, 64), (942, 963, 90), (966, 987, 115)]:
    for _row in range(_start, _end):
        _col_letter = chr(ord("D") + _row - _start)
        _ref_str = f"${_col_letter}${_anchor}"
        LicDsfConstraints.__annotations__[f"PV_Base!B{_row}"] = Literal[_ref_str]

# Language selector and lookup table (feed INDIRECT/VLOOKUP for language-dependent refs).
# START!L10 = VLOOKUP(K10, lookup!BB4:BC7, 2); evaluator does not support VLOOKUP, so L10 is constrained.
_LANG = Literal["English", "French", "Portuguese", "Spanish"]
_LANG_LOOKUP = Literal[
    "English", "French", "Portuguese", "Spanish", "Français", "Portugues", "Español"
]
LicDsfConstraints.__annotations__["START!L10"] = _LANG
LicDsfConstraints.__annotations__["START!K10"] = _LANG
for _r in range(4, 8):
    for _c in ("BB", "BC"):
        LicDsfConstraints.__annotations__[f"lookup!{_c}{_r}"] = _LANG_LOOKUP

LIC_DSF_CONSTRAINTS_DATA: dict[str, int | str | float] = {
    "PV_Base!A917": 64,
    "PV_Base!A941": 90,
    "PV_Base!A965": 115,
    **{
        f"PV_Base!A{r}": chr(ord("D") + r - _start)
        for _start, _end in [(918, 939), (942, 963), (966, 987)]
        for r in range(_start, _end)
    },
    **{
        f"PV_Base!B{r}": f"${chr(ord('D') + r - _start)}${_anchor}"
        for _start, _end, _anchor in [(918, 939, 64), (942, 963, 90), (966, 987, 115)]
        for r in range(_start, _end)
    },
    "START!L10": "English",
    "START!K10": "English",
    **{f"lookup!{c}{r}": "English" for r in range(4, 8) for c in ("BB", "BC")},
}


def get_dynamic_ref_config() -> DynamicRefConfig:
    """Return a DynamicRefConfig for constraint-based resolution of OFFSET/INDIRECT."""
    return DynamicRefConfig.from_constraints(
        LicDsfConstraints, LIC_DSF_CONSTRAINTS_DATA
    )
