# Benchmark baseline (run 1, 2026-03-10)

Run command:
```
uv run python -m src.lic_dsf_export
```

Timings:
- discover_targets: 0.00s
- build_graph: 457.02s
- populate_leaf_values: 175.90s
- enrich_graph: 330.71s
- generate_modules: 51.06s
- generate_setters_module: 264.21s

# Benchmark after optimization of workbook open operations (run 2, 2026-03-10)

Run command:
```
uv run python -m src.lic_dsf_export
```

Timings:
- discover_targets: 0.00s
- build_graph: 332.26s
- populate_leaf_values: 0.26s
- enrich_graph: 3.58s
- generate_modules: 29.14s
- generate_setters_module: 0.07s

# Benchmark after header iteration optimization per gh issue 12 (run 3, 2026-03-11)

Run command:
```
uv run python -m src.lic_dsf_export
```

Timings:
- discover_targets: 0.00s
- build_graph: 220.98s
- populate_leaf_values: 0.07s
- enrich_graph: 1.70s
- generate_modules: 25.79s
- generate_setters_module: 0.06s
