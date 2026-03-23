from __future__ import annotations

import re
import sys
import types

from src.configs import load_template_config


def test_2025_config_includes_figure_1_2_data_rows() -> None:
    # The template config imports excel_grapher for dynamic-ref wiring, but this
    # test only validates static export-range specs.
    excel_grapher_pkg = types.ModuleType("excel_grapher")
    excel_grapher_pkg.format_cell_key = lambda sheet, col, row: f"{sheet}!{col}{row}"
    excel_grapher_grapher = types.ModuleType("excel_grapher.grapher")
    excel_grapher_grapher.DynamicRefConfig = object
    excel_grapher_grapher.DependencyGraph = object
    sys.modules.setdefault("excel_grapher", excel_grapher_pkg)
    sys.modules["excel_grapher.grapher"] = excel_grapher_grapher

    cfg = load_template_config("2025-08-12")

    row_re = re.compile(r"^'Chart Data'!D(\d+):X\1$")
    exported_rows: set[int] = set()
    for entry in cfg.EXPORT_RANGES:
        m = row_re.match(entry["range_spec"])
        if m:
            exported_rows.add(int(m.group(1)))

    required_rows = {
        # Figure 1 chart data rows
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
        # Figure 2 rows not covered by existing stress-test blocks
        263,
        264,
        265,
        267,
        306,
        341,
        342,
        343,
    }

    missing = sorted(required_rows - exported_rows)
    assert not missing, f"Missing Figure 1/2 data rows in EXPORT_RANGES: {missing}"
