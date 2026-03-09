#!/usr/bin/env python3
"""
Generate RAG-based annotations for LIC-DSF indicator calculations.

This script builds a dependency graph for configured indicator rows, enriches nodes with
row/column labels, retrieves guidance-note context via a local embeddings collection, and
calls DeepSeek (OpenAI-compatible API) to produce concise annotations.

The embeddings collection is bootstrapped automatically (via `llm`'s Python API) if the
target collection does not already exist.
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import tempfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Literal
from urllib.request import urlopen

from dotenv import load_dotenv
import llm
from openai import AsyncOpenAI, OpenAI
import openpyxl
import openpyxl.utils.cell
import sqlite_utils

from excel_grapher.grapher import DependencyGraph, create_dependency_graph
from lic_dsf_config import (
    WORKBOOK_PATH,
    ensure_workbook_available,
    discover_targets_from_ranges,
    get_dynamic_ref_config,
)
from lic_dsf_labels import enrich_graph_with_labels, find_region_config


# Load environment variables from a local .env file (if present).
load_dotenv()


ANNOTATIONS_OUTPUT_PATH = Path("annotations.json")

GUIDANCE_COLLECTION_NAME = "lic-dsf-guidance"
GUIDANCE_EMBEDDING_MODEL = "text-embedding-3-small"
GUIDANCE_NOTE_PDF_URL = (
    "https://www.worldbank.org/content/dam/LIC%20DSF/Site%20File/assets/documentation/pp122617guidance-note-on-lic-dsf.pdf"
)
GUIDANCE_NOTE_PDF_PATH = Path("guidance_note/lic-dsf-guidance-note.pdf")
GUIDANCE_NOTE_TXT_PATH = Path("guidance_note/lic-dsf-guidance-note.txt")
GUIDANCE_CHUNK_CHARS = 1500
GUIDANCE_CHUNK_DIR = Path("lic-dsf-chunks")
GUIDANCE_CHUNK_PREFIX = "chunk_"
GUIDANCE_PROMPT_SNIPPET_CHARS = 1500


_embedding_collection: llm.Collection | None = None


def ensure_guidance_note_available() -> bool:
    """
    Ensure the LIC-DSF guidance note PDF exists locally.

    Downloads the PDF from `GUIDANCE_NOTE_PDF_URL` into `GUIDANCE_NOTE_PDF_PATH` if missing.
    """
    if GUIDANCE_NOTE_PDF_PATH.exists() and GUIDANCE_NOTE_PDF_PATH.stat().st_size > 0:
        return True

    GUIDANCE_NOTE_PDF_PATH.parent.mkdir(parents=True, exist_ok=True)

    try:
        with urlopen(GUIDANCE_NOTE_PDF_URL, timeout=60) as resp:
            with tempfile.NamedTemporaryFile(
                prefix=f".{GUIDANCE_NOTE_PDF_PATH.name}.",
                suffix=".download",
                dir=str(GUIDANCE_NOTE_PDF_PATH.parent),
                delete=False,
            ) as tmp:
                shutil.copyfileobj(resp, tmp)
                tmp_path = Path(tmp.name)

        if tmp_path.stat().st_size == 0:
            tmp_path.unlink(missing_ok=True)
            return False

        tmp_path.replace(GUIDANCE_NOTE_PDF_PATH)
        return True
    except Exception:
        return False


def _chunk_suffix(index: int, min_width: int = 2) -> str:
    """
    Produce a split(1)-style suffix: aa, ab, ..., az, ba, bb, ... (base-26 letters).
    """
    if index < 0:
        raise ValueError("index must be non-negative")
    digits: list[int] = []
    while True:
        digits.append(index % 26)
        index //= 26
        if index == 0:
            break
    while len(digits) < min_width:
        digits.append(0)
    return "".join(chr(ord("a") + d) for d in reversed(digits))


def _split_text_fixed_width(text: str, chunk_chars: int) -> list[str]:
    if chunk_chars <= 0:
        raise ValueError("chunk_chars must be positive")
    chunks: list[str] = []
    for i in range(0, len(text), chunk_chars):
        chunk = text[i : i + chunk_chars]
        if chunk.strip():
            chunks.append(chunk)
    return chunks


def _build_guidance_collection(db: sqlite_utils.Database, collection: llm.Collection) -> None:
    """
    Build the guidance-note embedding collection if it's missing.
    """
    if llm.Collection.exists(db, GUIDANCE_COLLECTION_NAME):
        return

    try:
        guidance_text = GUIDANCE_NOTE_TXT_PATH.read_text(encoding="utf-8")
    except FileNotFoundError:
        return

    chunks = _split_text_fixed_width(guidance_text, GUIDANCE_CHUNK_CHARS)
    entries: list[tuple[str, str]] = []
    for i, chunk in enumerate(chunks):
        chunk_id = f"{GUIDANCE_CHUNK_PREFIX}{_chunk_suffix(i)}"
        entries.append((chunk_id, chunk))

    # Keep a local copy of the chunk files for transparency/debugging.
    try:
        existing_chunk_files = any(GUIDANCE_CHUNK_DIR.glob(f"{GUIDANCE_CHUNK_PREFIX}*"))
        if not existing_chunk_files:
            GUIDANCE_CHUNK_DIR.mkdir(parents=True, exist_ok=True)
            for chunk_id, chunk in entries:
                (GUIDANCE_CHUNK_DIR / chunk_id).write_text(chunk, encoding="utf-8")
    except OSError:
        pass

    collection.embed_multi(entries, store=True, batch_size=100)


def get_embedding_collection() -> llm.Collection:
    """
    Get or create the LIC-DSF guidance note embedding collection.
    """
    global _embedding_collection
    if _embedding_collection is None:
        db_path = llm.user_dir() / "embeddings.db"
        db = sqlite_utils.Database(db_path)
        model = llm.get_embedding_model(GUIDANCE_EMBEDDING_MODEL)
        collection = llm.Collection(GUIDANCE_COLLECTION_NAME, db=db, model=model)
        _build_guidance_collection(db, collection)
        _embedding_collection = collection
    return _embedding_collection


def retrieve_guidance_context(query: str, n_results: int = 3) -> str:
    """
    Retrieve relevant context from the LIC-DSF guidance note via semantic search.
    """
    try:
        collection = get_embedding_collection()
        results = list(collection.similar(query, number=n_results))

        if not results:
            return "(No relevant context found in guidance note)"

        chunks: list[str] = []
        for r in results:
            content_raw = r.content or ""
            content = (
                content_raw[:GUIDANCE_PROMPT_SNIPPET_CHARS]
                if len(content_raw) > GUIDANCE_PROMPT_SNIPPET_CHARS
                else content_raw
            )
            chunks.append(f"[Score: {r.score:.3f}]\n{content}")

        return "\n\n---\n\n".join(chunks)
    except Exception as e:
        return f"(Error retrieving context: {e})"


def detect_annotation_axis(
    column_labels: list[str],
    row_labels: list[str],
) -> Literal["row", "column", "cell"]:
    """
    Auto-detect the annotation axis based on label patterns.
    """
    for label in column_labels:
        try:
            year = int(label)
            if 1900 <= year <= 2100:
                return "row"
        except ValueError:
            continue

    for label in row_labels:
        try:
            year = int(label)
            if 1900 <= year <= 2100:
                return "column"
        except ValueError:
            continue

    return "cell"


def get_annotation_key(
    sheet: str,
    row: int,
    col: int,
    axis: Literal["row", "column", "cell"],
) -> str:
    col_letter = openpyxl.utils.cell.get_column_letter(col)
    if axis == "row":
        return f"{sheet}!row{row}"
    if axis == "column":
        return f"{sheet}!col{col_letter}"
    return f"{sheet}!{col_letter}{row}"


def group_nodes_by_annotation_key(
    graph: DependencyGraph,
    enrichment_results: dict[str, dict[str, Any]],
) -> dict[str, list[str]]:
    """
    Group nodes by an annotation key for deduplication.
    """
    groups: dict[str, list[str]] = defaultdict(list)

    for node_key, data in enrichment_results.items():
        node = graph.get_node(node_key)
        if node is None:
            continue

        col_idx = openpyxl.utils.cell.column_index_from_string(node.column)

        region_config = find_region_config(node.sheet, node.row, col_idx)
        if region_config and "annotation_axis" in region_config:
            axis = region_config["annotation_axis"]
        else:
            axis = detect_annotation_axis(
                data.get("column_labels", []),
                data.get("row_labels", []),
            )

        annotation_key = get_annotation_key(node.sheet, node.row, col_idx, axis)
        groups[annotation_key].append(node_key)

    return dict(groups)


def get_node_summary(
    graph: DependencyGraph,
    node_key: str,
    enrichment_results: dict[str, dict[str, Any]],
) -> str:
    node = graph.get_node(node_key)
    if node is None:
        return node_key

    data = enrichment_results.get(node_key, {})
    row_labels = data.get("row_labels", [])
    col_labels = data.get("column_labels", [])

    labels: list[str] = []
    if row_labels:
        labels.extend(row_labels[:2])
    if col_labels:
        non_year_cols = [label for label in col_labels if not label.isdigit()]
        labels.extend(non_year_cols[:1])

    if labels:
        return f"{', '.join(labels)} ({node_key})"
    return node_key


def build_search_query(row_labels: list[str], column_labels: list[str]) -> str:
    non_year_labels: list[str] = []
    for label in column_labels:
        try:
            year = int(label)
            if 1900 <= year <= 2100:
                continue
        except ValueError:
            pass
        non_year_labels.append(label)

    all_labels = row_labels + non_year_labels
    return " ".join(all_labels[:5]) if all_labels else "debt sustainability indicator"


def build_annotation_prompt(
    row_labels: list[str],
    column_headers: list[str],
    sample_formula: str | None,
    parent_summaries: list[str],
    child_summaries: list[str],
    retrieved_context: str,
) -> str:
    if len(column_headers) > 5:
        col_str = ", ".join(column_headers[:5]) + f"... ({len(column_headers)} total)"
    else:
        col_str = ", ".join(column_headers) if column_headers else "(none)"

    parent_str = "\n".join(f"- {p}" for p in parent_summaries) if parent_summaries else "None"
    child_str = "\n".join(f"- {c}" for c in child_summaries) if child_summaries else "None"
    formula_str = sample_formula if sample_formula else "(no formula - input cell)"

    return f"""You are analyzing an Excel-based IMF/World Bank Low-Income Country Debt Sustainability Framework (LIC-DSF) model.

## Indicator Row
Row labels: {", ".join(row_labels) if row_labels else "(none)"}
Columns: {col_str}

## Formula
{formula_str}

## Inputs (parent cells feeding into this calculation)
{parent_str}

## Outputs (cells that depend on this value)
{child_str}

## LIC-DSF Guidance Note Context
{retrieved_context}

Based on the formula, inputs, outputs, and guidance note context, explain in 2-3 sentences:
1. What economic indicator or calculation this row represents
2. How the formula implements the DSF methodology (if applicable)
Keep the explanation concise and focused on the economic logic."""


def get_sample_formula(graph: DependencyGraph, node_keys: list[str]) -> str | None:
    for key in node_keys:
        node = graph.get_node(key)
        if node and node.formula:
            return node.formula
    return None


def get_parent_child_summaries(
    graph: DependencyGraph,
    node_keys: list[str],
    enrichment_results: dict[str, dict[str, Any]],
    max_each: int = 5,
) -> tuple[list[str], list[str]]:
    node_keys_set = set(node_keys)
    parent_keys: set[str] = set()
    child_keys: set[str] = set()

    for key in node_keys:
        for parent_key in graph.dependencies(key):
            if parent_key not in node_keys_set:
                parent_keys.add(parent_key)
        for child_key in graph.dependents(key):
            if child_key not in node_keys_set:
                child_keys.add(child_key)

    seen_parent_annotations: set[str] = set()
    unique_parents: list[str] = []
    for pk in parent_keys:
        node = graph.get_node(pk)
        if node is None:
            continue
        col_idx = openpyxl.utils.cell.column_index_from_string(node.column)
        data = enrichment_results.get(pk, {})
        axis = detect_annotation_axis(
            data.get("column_labels", []),
            data.get("row_labels", []),
        )
        ann_key = get_annotation_key(node.sheet, node.row, col_idx, axis)
        if ann_key not in seen_parent_annotations:
            seen_parent_annotations.add(ann_key)
            unique_parents.append(pk)

    seen_child_annotations: set[str] = set()
    unique_children: list[str] = []
    for ck in child_keys:
        node = graph.get_node(ck)
        if node is None:
            continue
        col_idx = openpyxl.utils.cell.column_index_from_string(node.column)
        data = enrichment_results.get(ck, {})
        axis = detect_annotation_axis(
            data.get("column_labels", []),
            data.get("row_labels", []),
        )
        ann_key = get_annotation_key(node.sheet, node.row, col_idx, axis)
        if ann_key not in seen_child_annotations:
            seen_child_annotations.add(ann_key)
            unique_children.append(ck)

    parent_summaries = [
        get_node_summary(graph, pk, enrichment_results) for pk in unique_parents[:max_each]
    ]
    child_summaries = [
        get_node_summary(graph, ck, enrichment_results) for ck in unique_children[:max_each]
    ]
    return parent_summaries, child_summaries


def get_deepseek_client() -> AsyncOpenAI:
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        raise ValueError("DEEPSEEK_API_KEY environment variable not set")
    return AsyncOpenAI(api_key=api_key, base_url="https://api.deepseek.com")


def get_deepseek_client_sync() -> OpenAI:
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        raise ValueError("DEEPSEEK_API_KEY environment variable not set")
    return OpenAI(api_key=api_key, base_url="https://api.deepseek.com")


async def annotate_node_group_async(
    client: AsyncOpenAI,
    graph: DependencyGraph,
    annotation_key: str,
    node_keys: list[str],
    enrichment_results: dict[str, dict[str, Any]],
    model: str = "deepseek-chat",
) -> tuple[str, str]:
    first_key = node_keys[0]
    first_data = enrichment_results.get(first_key, {})
    row_labels = first_data.get("row_labels", [])
    column_labels = first_data.get("column_labels", [])

    query = build_search_query(row_labels, column_labels)
    retrieved_context = retrieve_guidance_context(query)
    sample_formula = get_sample_formula(graph, node_keys)
    parent_summaries, child_summaries = get_parent_child_summaries(
        graph, node_keys, enrichment_results
    )

    prompt = build_annotation_prompt(
        row_labels=row_labels,
        column_headers=column_labels,
        sample_formula=sample_formula,
        parent_summaries=parent_summaries,
        child_summaries=child_summaries,
        retrieved_context=retrieved_context,
    )

    try:
        response = await client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert in IMF/World Bank debt sustainability analysis. Provide concise, accurate explanations.",
                },
                {"role": "user", "content": prompt},
            ],
            max_tokens=500,
            temperature=0.3,
        )
        annotation = response.choices[0].message.content or "(Empty response)"
        return annotation_key, annotation.strip()
    except Exception as e:
        return annotation_key, f"(Error generating annotation: {e})"


def annotate_node_group(
    graph: DependencyGraph,
    annotation_key: str,
    node_keys: list[str],
    enrichment_results: dict[str, dict[str, Any]],
    model: str = "deepseek-chat",
) -> str:
    first_key = node_keys[0]
    first_data = enrichment_results.get(first_key, {})
    row_labels = first_data.get("row_labels", [])
    column_labels = first_data.get("column_labels", [])

    query = build_search_query(row_labels, column_labels)
    retrieved_context = retrieve_guidance_context(query)
    sample_formula = get_sample_formula(graph, node_keys)
    parent_summaries, child_summaries = get_parent_child_summaries(
        graph, node_keys, enrichment_results
    )

    prompt = build_annotation_prompt(
        row_labels=row_labels,
        column_headers=column_labels,
        sample_formula=sample_formula,
        parent_summaries=parent_summaries,
        child_summaries=child_summaries,
        retrieved_context=retrieved_context,
    )

    try:
        client = get_deepseek_client_sync()
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert in IMF/World Bank debt sustainability analysis. Provide concise, accurate explanations.",
                },
                {"role": "user", "content": prompt},
            ],
            max_tokens=500,
            temperature=0.3,
            timeout=60,
        )
        annotation = response.choices[0].message.content or "(Empty response)"
        return annotation.strip()
    except Exception as e:
        return f"(Error generating annotation: {e})"


async def annotate_graph_async(
    graph: DependencyGraph,
    enrichment_results: dict[str, dict[str, Any]],
    max_groups: int | None = None,
    concurrency: int = 20,
    model: str = "deepseek-chat",
    verbose: bool = True,
) -> dict[str, str]:
    groups = group_nodes_by_annotation_key(graph, enrichment_results)

    if verbose:
        print(f"   Found {len(groups)} annotation groups")

    groups_to_process = list(groups.items())
    if max_groups:
        groups_to_process = groups_to_process[:max_groups]

    client = get_deepseek_client()
    semaphore = asyncio.Semaphore(concurrency)

    async def annotate_with_limit(
        key: str,
        nodes: list[str],
        index: int,
    ) -> tuple[str, str]:
        async with semaphore:
            if verbose:
                print(f"   [{index+1}/{len(groups_to_process)}] Annotating {key}...")
            return await annotate_node_group_async(
                client, graph, key, nodes, enrichment_results, model
            )

    tasks = [
        annotate_with_limit(key, nodes, i)
        for i, (key, nodes) in enumerate(groups_to_process)
    ]
    results = await asyncio.gather(*tasks)

    annotations: dict[str, str] = {}
    for annotation_key, annotation in results:
        annotations[annotation_key] = annotation
        for key in groups[annotation_key]:
            node = graph.get_node(key)
            if node:
                node.metadata["annotation"] = annotation
                node.metadata["annotation_key"] = annotation_key

    return annotations


def main() -> None:
    print("=" * 70)
    print("LIC-DSF Indicator Annotation (RAG + DeepSeek)")
    print("=" * 70)

    if not ensure_guidance_note_available() and not GUIDANCE_NOTE_TXT_PATH.exists():
        print("Error: Guidance note is not available locally.")
        print(f"Expected PDF at: {GUIDANCE_NOTE_PDF_PATH}")
        print(f"Expected text at: {GUIDANCE_NOTE_TXT_PATH}")
        return

    if not ensure_workbook_available(WORKBOOK_PATH):
        print(f"Error: Workbook not available at {WORKBOOK_PATH}")
        return

    print("\n1. Collecting target cells from configured ranges...")
    all_targets = discover_targets_from_ranges(WORKBOOK_PATH)

    print(f"   Total targets: {len(all_targets)}")
    if not all_targets:
        print("No target cells found. Exiting.")
        return

    print("\n2. Building dependency graph...")
    graph = create_dependency_graph(
        WORKBOOK_PATH,
        all_targets,
        load_values=False,
        max_depth=50,
        dynamic_refs=get_dynamic_ref_config(),
        use_cached_dynamic_refs=False,
    )
    print(f"   Nodes in graph: {len(graph)}")

    print("\n3. Enriching nodes with row/column labels...")
    enrichment_results = enrich_graph_with_labels(graph, WORKBOOK_PATH)
    print(f"   Enriched nodes: {len(enrichment_results)}")

    print("\n4. Annotating node groups...")
    annotations = asyncio.run(
        annotate_graph_async(
            graph,
            enrichment_results,
            concurrency=20,
            model="deepseek-chat",
            verbose=True,
        )
    )

    payload = {
        "workbook": str(WORKBOOK_PATH),
        "collection": GUIDANCE_COLLECTION_NAME,
        "model": "deepseek-chat",
        "annotations": annotations,
    }
    ANNOTATIONS_OUTPUT_PATH.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"\n5. Wrote {len(annotations)} annotations to {ANNOTATIONS_OUTPUT_PATH}")


if __name__ == "__main__":
    main()

