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
- `map_lic_dsf_indicators.py`
  - Dependency mapping + label enrichment + `enrichment_audit.json` export
- `lic_dsf_annotate.py`
  - RAG + DeepSeek annotations + `annotations.json` export

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
uv run python map_lic_dsf_indicators.py
```

### Inputs

- `workbooks/lic-dsf-template.xlsm` (default path configured via `WORKBOOK_PATH`)
- Indicator-row configuration in `INDICATOR_CONFIG` inside `map_lic_dsf_indicators.py`
- Label extraction configuration in `REGION_CONFIG` inside `map_lic_dsf_indicators.py`

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

- Workbook: `workbooks/lic-dsf-template.xlsm` (imported from `map_lic_dsf_indicators.WORKBOOK_PATH`)
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

