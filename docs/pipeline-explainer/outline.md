The problem:
Excel workbooks are powerful — but the logic runs *inside the application*. We want the *same logic* in Python so economists can run scenarios without opening Excel.

Manual Excel work is *laborious*, *time-consuming*, and *error-prone*. Workflows cannot be *typed*, *tested*, or even *audited*. It can't be *automated* or done *at scale*.

The solution in one sentence:
We *trace* what each cell depends on, *translate* the formulas to Python, and *ship a library* that runs the same math.

But that's a big challenge. The LIC DSF:

[Cycle through bullets one per slide, with transition animation]
- has over *800,000 cells*, including *200,000 formula cells*
- calls *55 unique functions*, nested up to *7 levels* deep
- uses lookups and offsets that are *tightly coupled to Excel's data model*

Option 1:
*Drive Excel from Python*. Pros: modest speedup, familiar tooling, correctness guarantees. Cons: most of the limitations of Excel still apply.

Option 2:
*Have AI agents translate to Python*. Pros: leverages the "bitter lesson," improves as AI models do. Cons: unpredictable output shape, no correctness guarantees, unreproducible for new template versions.

Option 3:
*Translate to Python mechanically*. Pros: predictable output shape, correctness guarantees, reproducible for new template versions. Cons: hard to get right, output is still Excel-shaped.

AI agents help make Option 3 tractable. And we can *post-process the output to make it more Python-shaped*.

The lesson:
For interpretability, reproducibility, and correctness guarantees, *don't AI-generate the code*. *AI-generate the code that generates the code*.

Architecture:
```mermaid
flowchart TB
  A@{ shape: brace-r, label: "Core engine to trace & translate Excel to Python" }
  EG@{ shape: lin-cyl, label: "excel-grapher" }
  B@{ shape: brace-r, label: "Configuration and post-processing to make output Python-shaped" }
  LIC@{ shape: lin-cyl, label: "lic-dsf-programmatic-extraction" }
  C@{ shape: brace-r, label: "The user-facing library\n(our deliverable)" }
  PY@{ shape: rectangle, label: "py-lic-dsf" }
  
  A ~~~ B ~~~ C
  EG --> LIC --> PY
```

[Zoom effect to excel-grapher if possible (zoom to upper-third of screen with fade and replace, maybe?)]

Let's start with the core engine. How do we mechanically trace & translate Excel to Python?
```mermaid
flowchart LR
  A@{ shape: brace-r, label: "Core engine to trace & translate Excel to Python" }
  EG@{ shape: lin-cyl, label: "excel-grapher" }
  A ~~~ EG
```

The Excel template is architected a big-ass data object (BADO) plus a stable engine that knows how to interpret and compute that object.

```mermaid
flowchart LR
    A@{ img: "/home/chriscarrollsmith/Documents/Software/Excel_Extraction/lic-dsf-programmatic-extraction/docs/pipeline-explainer/excel-xlsm-icon.svg", label: "Excel workbook", h: 80, constraint: "on" }
    B@{ img: "/home/chriscarrollsmith/Documents/Software/Excel_Extraction/lic-dsf-programmatic-extraction/docs/pipeline-explainer/excel-app-icon.svg", label: "Excel application", h: 80, constraint: "on" }
    A ~~~ B
```

It's mostly a solved problem to replicate this BADO + engine approach in Python.

```mermaid
flowchart TB
    A@{ img: "/home/chriscarrollsmith/Documents/Software/Excel_Extraction/lic-dsf-programmatic-extraction/docs/pipeline-explainer/excel-xlsm-icon.svg", label: "Excel workbook", h: 80, constraint: "on" }
    B@{ img: "/home/chriscarrollsmith/Documents/Software/Excel_Extraction/lic-dsf-programmatic-extraction/docs/pipeline-explainer/excel-app-icon.svg", label: "Excel application", h: 80, constraint: "on" }
    C@{ shape: cyl, label: "Python dictionary" }
    D@{ shape: lin-rect, label: "Excel emulator"}
    A --> C
    B --> D
```

The workbook becomes a *Python dictionary* where every *cell* is an entry with its *formula* and *value* keyed to the *address* of the cell:

```python
xlsm_contents = {
    "Sheet1!A1": {
        "formula": "=B1+C1",
        "value": 5
    }
}
```

The *engine* gets implemented as what's called an *abstract syntax tree parser* (AST parser): a small program that can break down an Excel formula into a branching tree of operations and translate each operation into a Python equivalent. Thankfully, we can borrow from smart people like the creator of the *formulas* library who already mostly built this part!

Simple example AST:
```mermaid
flowchart TB
    A1["A1"]
    Add["+"]
    ref_B1["B1"]
    ref_C1["C1"]
    ref_B1 --> Add
    ref_C1 --> Add
    Add --> A1
```

The hard parts of implementing BADO + engine:
- If we only want particular output cells, how do we limit our extraction to only the relevant inputs?
- What sequence do you compute the cells in?
We solve these problems by building a *graph* of relationships between cells.

For any cell we care about (e.g. a stress-test result), we ask: *which other cells does its formula use?* We repeat until we hit numbers or text. That gives us a *dependency graph*.

We call this *target-driven dependency tracing*. Give `excel-grapher` a target cell; it follows the formula references until it reaches leaves. The result is a graph of *nodes (cells)* and *edges (depends-on)*.

Smallest example: the two-cell demo:

| A1 | B1 |
|----|----|
| 10 | =A1 × 2 |

B1 depends on A1. The graph is *two nodes* and *one edge*. If B1 is our *target*, A1 is the *leaf*. We need to extract A1 to compute B1, but we don't need C1 (or any other cells).

```mermaid
flowchart LR
  A1["A1 = 10"]
  B1["B1 = A1 × 2"]
  B1 --> A1
```

BADO + evaluator is the first layer of excel-grapher. *It already works today.* We can run the workbook this way and get results.

```mermaid
flowchart LR
    excel_grapher["excel-grapher"]
    subgraph layer1["Layer 1: BADO + engine"]
        grapher["grapher"]
        evaluator["evaluator"]
    end
    excel_grapher --- layer1
```

Limitations of the first layer:
- *Poor separation of concerns* — formulas live in the data layer, so data and economic logic are mixed in one big structure.
- *Non-transparent economic logic* — the model consists of 200,000 atomic operations; neither readable nor maintainable.

What we want in a Python library: the *data layer* holds constants and inputs; *computational logic* lives as code.

That brings us to the second layer: the *exporter* module.

```mermaid
flowchart LR
    excel_grapher["excel-grapher"]
    subgraph layer1["Layer 1: BADO + engine"]
        grapher["grapher"]
        evaluator["evaluator"]
    end
    subgraph layer2["Layer 2"]
        exporter
    end
    excel_grapher --- layer1
    excel_grapher --- layer2
```

The exporter turns formula cells into *Python functions* and outputs them, along with relevant parts of the engine, as a standalone Python library. Non-formula cells stay as a Python dictionary. That gives us a clear split: *data* in the dictionary, *logic* in code.

The exporter is *configurable*:
- Specify *targets* (which outputs you want) → it generates functions that produce those outputs.
- Specify how to *group inputs* → it generates setters for those groups.
- Mark cells as *constants* (used in the computation but not user-settable) so they stay out of the public API.

Configuration is where *domain knowledge* belongs: which cells are inputs, which are outputs, which are internal constants. For the LIC DSF template, that configuration lives in the `lic-dsf-programmatic-extraction` repository.

Where we're headed:
- Today we still have *one function per formula cell*.
- Goal: *one function per economic model concept*.
- We'll use *static analysis* and *graph analysis* to find groups of functions that should be collapsed. Most of that post-processing will live in `lic-dsf-programmatic-extraction`; some of it may be reusable and live in `excel-grapher`.
