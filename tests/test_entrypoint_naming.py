from __future__ import annotations

import json
import sys
import types


# src.lic_dsf_export imports excel_grapher at module import time; stub it for tests.
excel_grapher_pkg = types.ModuleType("excel_grapher")
excel_grapher_pkg.format_cell_key = lambda sheet, col, row: f"{sheet}!{col}{row}"
excel_grapher_exporter = types.ModuleType("excel_grapher.exporter")
excel_grapher_exporter.CodeGenerator = object
excel_grapher_grapher = types.ModuleType("excel_grapher.grapher")
excel_grapher_grapher.get_calc_settings = lambda *args, **kwargs: None
excel_grapher_grapher.DependencyGraph = object
excel_grapher_grapher.Node = object
excel_grapher_grapher.create_dependency_graph = lambda *args, **kwargs: None
excel_grapher_grapher.DynamicRefConfig = object
sys.modules.setdefault("excel_grapher", excel_grapher_pkg)
sys.modules["excel_grapher.exporter"] = excel_grapher_exporter
sys.modules["excel_grapher.grapher"] = excel_grapher_grapher

from src.lic_dsf_export import build_entrypoints


def test_build_entrypoints_omits_sheet_prefixes(tmp_path) -> None:
    audit = {
        "by_sheet": {
            "Output 2-1 Stress_Charts_Ex": {
                "cells": [{"address": "D239", "row_labels": ["Baseline"]}]
            },
            "Output 2-2 Stress_Charts_Pub": {
                "cells": [{"address": "D239", "row_labels": ["Baseline"]}]
            },
            "Chart Data": {
                "cells": [{"address": "D10", "row_labels": ["Overall rating"]}]
            },
        }
    }
    audit_path = tmp_path / "enrichment_audit.json"
    audit_path.write_text(json.dumps(audit), encoding="utf-8")

    targets = [
        "Output 2-1 Stress_Charts_Ex!D239",
        "Output 2-2 Stress_Charts_Pub!D239",
        "Chart Data!D10",
    ]
    export_ranges = [{"range_spec": "'Chart Data'!D10:D10", "entrypoint_mode": "per_cell"}]

    entrypoints = build_entrypoints(targets, audit_path, export_ranges)

    assert "baseline" in entrypoints
    assert "baseline_2" in entrypoints
    assert "overall_rating_d" in entrypoints
    assert "output_2_1_stress_charts_ex_baseline" not in entrypoints
    assert "chart_data_overall_rating_d" not in entrypoints
