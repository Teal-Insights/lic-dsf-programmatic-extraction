## LIC-DSF programmatic extraction

This repo contains scripts to analyze an IMF/World Bank LIC-DSF Excel workbook:

- **Dependency mapping**: identify formula cells in configured indicator rows, build a dependency
  graph, and enrich nodes with human-readable row/column labels.
- **Code generation**: export workbook formulas as a standalone Python package that can be published
  to PyPI and used without Excel.
- **RAG-based annotation**: retrieve relevant context from the LIC-DSF guidance note using a local
  embeddings collection, then call DeepSeek (`deepseek-chat`) to generate short annotations for
  indicator groups.

## Template versioning

The World Bank periodically releases new LIC-DSF template workbooks. Each template version can
differ in structure (sheet layout, cell ranges, formulas), so all template-specific configuration
lives in its own directory under `src/configs/<date>/`:

```
src/configs/
  2025-08-12/
    config.py          # workbook path, export ranges, constraints, region config, etc.
    input_groups.json   # generated artifact
    enrichment_audit.json
  2026-01-31/
    config.py
    input_groups.json
    enrichment_audit.json
```

Each template version produces an **independent PyPI package** (e.g. `lic-dsf-2026-01-31`) so that
users on different template versions can coexist. When a new template is released:

1. Add the workbook to `workbooks/`
2. Create `src/configs/<date>/config.py` (copy the most recent config and adjust)
3. Run the pipeline with `--template <date>`
4. Test and publish the generated package

## Repository layout

- `workbooks/` — source-of-truth workbooks (one per template version)
- `src/configs/<date>/config.py` — per-template configuration (ranges, constraints, region config)
- `src/configs/<date>/*.json` — per-template generated artifacts
- `src/lic_dsf_config.py` — shared type definitions and utility functions
- `src/lic_dsf_pipeline.py` — shared graph + classification utilities
- `src/lic_dsf_labels.py` — label extraction helpers
- `src/lic_dsf_export.py` — code generation + enrichment audit
- `src/lic_dsf_group_inputs.py` — input grouping + `input_groups.json` export
- `src/lic_dsf_input_setters.py` — shared setter helpers used by generated export package
- `src/lic_dsf_annotate.py` — DeepSeek annotations
- `guidance_note/` — LIC-DSF guidance note PDF and text
- `dist/lic-dsf-<date>/` — generated Python packages (one per template)

## Prerequisites

- Python version per `pyproject.toml`
- Dependencies installed via `uv`
- A DeepSeek API key for annotation runs

## Setup

Create a virtual environment and install deps:

```bash
uv sync
```

Set your DeepSeek key (used by `src/lic_dsf_annotate.py`):

```bash
export DEEPSEEK_API_KEY="..."
```

Optionally, store it in a `.env` file (loaded by `src/lic_dsf_annotate.py`):

```bash
DEEPSEEK_API_KEY=...
```

## Pipeline scripts

All scripts require a `--template` argument specifying which template version to use. Available
templates are auto-discovered from `src/configs/`.

### Script 1: Dependency mapping + enrichment audit

Builds a dependency graph, enriches nodes with row/column labels, and writes an audit JSON.

```bash
uv run python -m src.lic_dsf_export --template 2026-01-31 --audit-only
```

**Inputs**: workbook and configuration from `src/configs/2026-01-31/config.py`

**Output**: `src/configs/2026-01-31/enrichment_audit.json` (overwritten on every run)

### Script 2: Export formulas to standalone Python code

Discovers targets, builds a dependency graph, and uses `excel-grapher`'s `CodeGenerator` to emit a
standalone Python package.

```bash
uv run python -m src.lic_dsf_export --template 2026-01-31
```

**Output**: `dist/lic-dsf-2026-01-31/lic_dsf_2026_01_31/` (overwritten on every run)

### Script 3: Group inputs for setter generation

Groups hardcoded input cells into semantically labeled clusters for setter code generation.

```bash
uv run python -m src.lic_dsf_group_inputs --template 2026-01-31
```

**Output**: `src/configs/2026-01-31/input_groups.json` (overwritten on every run)

### Script 4: RAG-based annotation (Guidance Note + DeepSeek)

Retrieves guidance-note context via embeddings and calls DeepSeek to generate concise annotations.

```bash
uv run python -m src.lic_dsf_annotate --template 2026-01-31
```

**Inputs**: workbook, guidance note text (`guidance_note/lic-dsf-guidance-note.txt`), `DEEPSEEK_API_KEY`

**Output**: `src/configs/2026-01-31/annotations.json` (overwritten on every run)

## Recommended sequence

```bash
# 1. (Optional) Generate enrichment audit
uv run python -m src.lic_dsf_export --template 2026-01-31 --audit-only

# 2. (Optional) Generate input groups for setters
uv run python -m src.lic_dsf_group_inputs --template 2026-01-31

# 3. (Optional) Generate annotations
uv run python -m src.lic_dsf_annotate --template 2026-01-31

# 4. Core export step — generates the Python package
uv run python -m src.lic_dsf_export --template 2026-01-31
```

## Using generated input setters

The generated package exposes a context object with helper setters derived from `input_groups.json`.

- **Year-series setters**: accept `{year: value}` (primary) and also `values + start_year` (secondary).
- **Range setters (scalars / 1D / 2D tables)**: accept a scalar, 1D sequence, or 2D sequence-of-sequences matching the range shape.

Example:

```python
import lic_dsf_2026_01_31 as lic_dsf

ctx = lic_dsf.make_context()

# Year-series: dict form (recommended)
assignment = ctx.set_ext_debt_data_external_debt_excluding_locally_issued_debt({2023: 123, 2026: None})

# 1D range
ctx.set_ext_debt_data_ida_new_60_year_credits([1] * 14)

# Load all inputs from a filled-out template (requires optional fastpyxl)
ctx.load_inputs_from_workbook("workbooks/lic-dsf-template-2026-01-31.xlsm")
```

## Embeddings store (how it works)

Semantic search uses the [`llm`](https://llm.datasette.io/) library's embeddings database:

- DB location: `~/.config/io.datasette.llm/embeddings.db`
- Collection name: `lic-dsf-guidance`
- Embedding model: `text-embedding-3-small`

When bootstrapping, `src/lic_dsf_annotate.py`:

- Splits the guidance note text into ~1500-character chunks
- Stores embeddings for those chunks in the `lic-dsf-guidance` collection
- Optionally writes chunk files under `lic-dsf-chunks/` if none are present

### Resetting / rebuilding embeddings

If you need to force a rebuild, delete the collection using the `llm collections` entrypoint:

```bash
uv run llm collections list
uv run llm collections delete lic-dsf-guidance
```

Then rerun:

```bash
uv run python -m src.lic_dsf_annotate --template 2026-01-31
```
