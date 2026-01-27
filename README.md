## LIC-DSF programmatic extraction

This repo contains scripts to analyze an IMF/World Bank LIC-DSF Excel workbook:

- **Dependency mapping**: identify formula cells in configured indicator rows, build a dependency
  graph, and enrich nodes with human-readable row/column labels.
- **RAG-based annotation**: retrieve relevant context from the LIC-DSF guidance note using a local
  embeddings collection, then call DeepSeek (`deepseek-chat`) to generate short annotations for
  indicator groups.

The code is intentionally split into two scripts so “mapping” stays mapping-only, and the LLM/RAG
workflow lives in a dedicated entrypoint.

## Repository layout

- `workbooks/`
  - `lic-dsf-template.xlsm`: **source-of-truth** workbook for this workflow
  - (optional) other workbooks for comparison/testing
- `guidance_note/`
  - `lic-dsf-guidance-note.pdf`: guidance note PDF (downloaded automatically if missing)
  - `lic-dsf-guidance-note.txt`: plaintext guidance note used for semantic search
- `lic-dsf-chunks/`
  - Local chunk files used to build the embeddings store (created only if missing)
- `lic_dsf_annotate.py`
  - DeepSeek annotations + `annotations.json` export
- `lic_dsf_pipeline.py`
  - Shared graph + classification utilities used by export and input grouping
- `lic_dsf_labels.py`
  - Workbook configuration and label extraction helpers
- `lic_dsf_group_inputs.py`
  - Input grouping + `input_groups.json` export (inputs only, constants filtered)
- `lic_dsf_input_setters.py`
  - Shared setter helpers used by generated export package

## Prerequisites

- Python version per `pyproject.toml`
- Dependencies installed via `uv`
- A DeepSeek API key for annotation runs

## Setup

Create a virtual environment and install deps:

```bash
uv sync
```

Set your DeepSeek key (used by `lic_dsf_annotate.py`):

```bash
export DEEPSEEK_API_KEY="..."
```

Optionally, store it in a `.env` file (loaded by `lic_dsf_annotate.py`):

```bash
DEEPSEEK_API_KEY=...
```

## Script 1: Dependency mapping + enrichment audit

Runs dependency discovery for the configured indicator rows, builds a dependency graph, enriches
nodes with row/column labels, and writes an audit JSON.

### Run

```bash
uv run python lic_dsf_export.py --audit-only
```

### Inputs

- `workbooks/lic-dsf-template.xlsm` (default path configured via `WORKBOOK_PATH`)
- Indicator-row configuration in `INDICATOR_CONFIG` inside `lic_dsf_labels.py`
- Label extraction configuration in `REGION_CONFIG` inside `lic_dsf_labels.py`

### Output

- `enrichment_audit.json`
  - Contains summary statistics + sheet-by-sheet details for extracted labels.

### Rerun behavior

- `enrichment_audit.json` is **overwritten** on every run.

## Script 2: RAG-based annotation (Guidance Note + DeepSeek)

This script:

- Finds formula cells in the configured indicator rows
- Builds a dependency graph
- Enriches nodes with row/column labels (re-uses mapping logic)
- Retrieves relevant guidance-note context using an embeddings collection
- Calls DeepSeek (`deepseek-chat`) to generate concise annotations

### Run

```bash
uv run python lic_dsf_annotate.py
```

### Inputs

- Workbook: `workbooks/lic-dsf-template.xlsm` (imported from `lic_dsf_labels.WORKBOOK_PATH`)
- Guidance note text: `guidance_note/lic-dsf-guidance-note.txt`
- DeepSeek API key: `DEEPSEEK_API_KEY`

### Output

- `annotations.json`
  - Includes workbook path, embedding collection name, model name, and a map of annotation keys to
    generated text.

### Rerun behavior

- `annotations.json` is **overwritten** on every run.
- The embeddings collection is **not** rebuilt on every run:
  - If the collection exists, it is reused.
  - If missing, `lic_dsf_annotate.py` will bootstrap it from the guidance note.

## Embeddings store (how it works)

Semantic search uses the [`llm`](https://llm.datasette.io/) library’s embeddings database:

- DB location: `~/.config/io.datasette.llm/embeddings.db`
- Collection name: `lic-dsf-guidance`
- Embedding model: `text-embedding-3-small`

When bootstrapping, `lic_dsf_annotate.py`:

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
uv run python lic_dsf_annotate.py
```

## Script 3: Export formulas to standalone Python code

This script:

- Discovers formula targets from `INDICATOR_CONFIG`
- Builds a dependency graph
- Uses `excel-formula-expander`'s `CodeGenerator` to emit a small Python package

### Run

```bash
uv run python lic_dsf_export.py
```

### Audit-only mode

```bash
uv run python lic_dsf_export.py --audit-only
```

### Output

- `export/<normalized-workbook-stem>/` (overwritten on every run)

### Using generated input setters

The generated package exposes a context object with helper setters derived from `input_groups_export.json`.

- **Year-series setters**: accept `{year: value}` (primary) and also `values + start_year` (secondary).
- **Range setters (scalars / 1D / 2D tables)**: accept a scalar, 1D sequence, or 2D sequence-of-sequences matching the range shape.

Example:

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path("export").resolve()))
import lic_dsf

ctx = lic_dsf.make_context()

# Year-series: dict form (recommended)
assignment = ctx.set_ext_debt_data_external_debt_excluding_locally_issued_debt({2023: 123, 2026: None})

# 1D range
ctx.set_ext_debt_data_ida_new_60_year_credits([1] * 14)
```

## Script 4: Group inputs for setter generation

This script:

- Discovers formula targets from `INDICATOR_CONFIG`
- Builds a dependency graph
- Populates leaf values and classifies constants vs inputs
- Enriches input cells with row/column labels
- Groups inputs into labeled clusters and writes JSON

### Run

```bash
uv run python lic_dsf_group_inputs.py
```

### Output

- `input_groups.json` (overwritten on every run)

### Export integration

If you copy or rename the output to `input_groups_export.json`, `lic_dsf_export.py` will
generate setters in the exported package using those groups.

## Recommended sequence

1. `lic_dsf_export.py --audit-only` (optional, if you want updated `enrichment_audit.json`)
2. `lic_dsf_group_inputs.py` (optional, if you want updated input groups for setters)
3. `lic_dsf_annotate.py` (optional today; planned to inform export docstrings)
4. `lic_dsf_export.py` (core export step)
