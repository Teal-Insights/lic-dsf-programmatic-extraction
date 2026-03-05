# Rough Notes: Proposed LIC DSF External Debt Risk Rating Python package work plan

Owner: Christopher Smith
Date: February 3, 2026
Status: Reference
File Type: Rough Notes
Summary: The project is divided into two phases: Phase 1 focuses on building Excel extraction tools, while Phase 2 aims to create a user-friendly Python library for economists. Key deliverables include user research on the LIC DSF, scaling the library to include all stress tests, improving API usability, implementing a verification framework, and enhancing code conciseness. Future phases will extend functionality to support specific workflows and provide analyses of data flow within the library.
Created by: Christopher Smith
Created: February 3, 2026 8:48 AM
Last Edited Time: February 10, 2026 11:21 AM

NOTE: This document is 100% human-written.

# Introduction

I’ve been mentally modeling this project in roughly two phases. Phase 1 is drawing to a close, while Phase 2 is in its very early stages.

1. **Phase 1: Building abstract, generalizable Excel extraction and translation tooling** (mainly with an eye to extracting and translating the LIC DSF, but with a secondary goal of supporting similar projects later).
2. **Phase 2: Shaping the output into an economist-friendly Python library**, with economic model-shaped functions, a well-organized and understandable API, and robust documentation.

We’re currently transitioning from Phase 1 to Phase 2.

There remains some Phase 1 work to be done, especially increasing test robustness. And some Phase 2 work will be implemented as reusable modules in the Excel extraction and codegen layers.

But the fundamentals of our Phase 1 implementation seem sound, so I’m pivoting to doing more Phase 2 work on the economist-facing layer.

# Phase 1 mop-up: Increase test robustness

- Verify that 100% pass rate with rtol 1e-5 is rigorously enforced throughout the test suite (configure centrally in `conftest.py`)
- Do more input manipulation, especially to flip zero values to non-zero values in the golden master tests
- Flip conditional toggles to test more execution pathways through the notebook
- Add better unit test coverage for formula implementations, especially unhappy path testing
- Come up with a better verification strategy for dependency detection in `excel-grapher` (good suggestion from Kweku)

# Phase 2: Reshape codegen outputs into an ergonomic user-facing library

In this phase, it would be highly advantageous to regularly have the team’s eyes on the work, especially the economic domain experts (Teal and Reuben).

At a high level, I see five near-term deliverables:

## Deliverable 1: LIC DSF user research

Because I’m pivoting my focus to making this ergonomic for macroeconomists to use, I want to finish the LIC DSF EdX course and also read the LIC DSF guidance note before I go much further. That will give me necessary insight in use cases and user stories.

**Concrete deliverable:** 3-7 pages of notes on how to mentally model the LIC DSF for the purpose of shaping the library API.

**Timeframe:** 2-7 days

## Deliverable 2: Scale to the other tabs

The completed Python library MVP from Phase 1 implements three (3) of the six (6) stress tests that contribute the external debt risk rating. We need to scale our implementation to include at least the baseline, the remaining three stress test tabs, the debt thresholds, and the external debt risk rating. Deliverable 1 should inform a decision as to whether any other targets should be added to this list (and we should also discuss this as a team). Likely will require implementing and unit testing some more Excel functions in Python.

**Concrete deliverable:** Published py-lic-dsf version that includes `compute_*` functions for all six stress test tabs, the baseline, the thresholds, and the final external debt risk rating.

**Timeframe:** 2-7 days

## Deliverable 3: Make the library API interpretable/usable

I currently have:

1. AI summarization of each cell’s economic logic based on the guidance note (currently not actually used in the library), and
2. Sheet name and row/column label detection for each cell (used for naming functions and grouping outputs into time series).

This helps, but falls well short of supporting real usability.

1. **Problem:** There are ~200 input-setter methods on `LicDsfContext`. We have no system for organizing them.
**Solution:** Organize methods into namespaced subclasses, perhaps by sheet name, hierarchical label, and/or some kind of use-case classification system informed by the guidance note (e.g., HistoricalData, ProjectionData, WorkbookConfiguration, etc.).

```python
class LicDsfContext:
    def __init__(self):
        self.historical_data = self.HistoricalData(self)
    
    class HistoricalData:
        def __init__(self, context):
            self.context = context
        
        def set_some_time_series(self):
            # Now you can access self.context for outer instance
            return "Time series set!"

ctx = LicDsfContext()
ctx.historical_data.set_some_time_series()
```

[Python classes and their application in py-lic-dsf](https://www.loom.com/share/d3f178493cbb44d08e5b6404dffd32f8)

1. **Problem:** We have no system for detecting documenting which inputs are necessary to compute which outputs.
**Solution:** We can *detect* which inputs contribute to which outputs by traversing the subgraph for that set of outputs. We have to do it in our extraction pipeline and export this as library documentation rather than do it in the library at runtime, because the library doesn’t have access to the graph.
2. **Problem:** We have no system to surface particular inputs that are more commonly used than others.
**Solution:** We should probably have two documentation layers:
    1. a README that covers the most common use cases and surfaces the most critical inputs, and
    2. comprehensive documentation for all methods exposed by the library.
    
    Also, we should:
    
    1. organize methods into subclasses as described above under Problem 1, and
    2. sequence the methods within each class according to some importance heuristic, perhaps driven by programmatic AI analysis of the guidance note.
    
    We should also explore having people fill out parameterization sheets first, then have helpers to tell them which other inputs they need to fill out given the configuration they selected.
    
3. **Problem:** We have no system for populating historical data from public datasets (as I think some of the macros in the new LIC DSF template are designed to do).
**Solution:** We should investigate and perhaps translate any workbook macros involved in populating historical data. If I’m wrong about the macros, then we could explore rolling our own data-API-calling methods.
4. **Problem:** We have no explanatory docstrings in the code, and no comprehensive documentation for user-facing methods.
**Solution:** We can use the AI annotations to populate docstrings at least for the `compute_*` methods, if not the [`internals.py`](http://internals.py) functions. Also, we should programmatically generate comprehensive documentation for every input-setting method by using AI + RAG + graph traversal to identify how that set of inputs contributes to the target outputs.
5. **Problem:** We have no way, currently, of detecting and collecting hierarchical row labels (such as appear on the `Ext_Debt_Data` tab) where the parent labels appear in previous rows and the hierarchical relationship is indicated by formatting cues like bold text and indentation.
**Solution:** This is probably programmatically detectable with heuristics, but if not, then manual/agentic configuration of the label detection workflow might be required.

![Drawing.sketchpad.jpeg](work-plan-files/Drawing.sketchpad.jpeg)

1. **Problem:** Some function names are extremely long. (Currently I am setting the method name to sheet name + all collected non-year labels, which blows up method name lengths for nested labels and long labels.)
**Solution:** We need some kind of name-shortening heuristic. Maybe an AI model that does rules-based shortening for names above a certain length (and then we cache the generated names so they stay the same next time we regenerate).

**Concrete deliverables:** We’ll need at least one PR implementing the solution for each of the seven problems listed above.

**Timeframe:** Honestly, it’s a big job. Likely 2-4 weeks to solve all seven problems, so we might want to implement selectively and prioritize the most important ones.

## Deliverable 4: Verification framework

We need to implement comprehensive golden master tests inside the public `py-lic-dsf` repository.

**Concrete deliverables:** Golden master test file that applies the patterns developed in the `excel-formula-expander` layer but applies them to the generated code. Should apply lessons from the Phase 1 mop-up operation above.

## Deliverable 5: Conciseness/modularization of internals

We’ve got something like 400,000 lines of code right now. And that code is “Excel-shaped”: it makes use of Excel addresses, implements a separate function for every formula cell in the dependency tree, exports a dictionary entry for every hardcoded cell in the dependency tree, and uses an internal “Excel runtime” with implementations of Excel formulas.

Ideally, we want the final library to be both: 1. a lot more concise, and 2. a lot more “economic model-shaped,” with function names and module boundaries that correspond to economic concepts. And we want this for the internals as well as the inputs/outputs.

I propose three strategies:

### **1. Graph compression**

Say we have a formula graph that includes this chain of formula references, where `B1` *only* has edges with `A1` and `C1`, and with no other cells: `A1 → B1 → C1`. The graph can be shortened by eliminating `B1` by substitution, so the new graph is `A1 -> C1`. This is an easy case, but I think there are known algorithms for graph compression that are more sophisticated and can give bigger compression gains.

### 2. Graph segmentation

By analyzing the graph, we can identify sections of the graph that are “self-contained”. For example, say we have three cells, A1, B1, and C1, that all point to E1. E1, in turn, is the parent of a tree that’s entirely self contained (ends in leaf nodes) except for one cell (say, J1) that’s referenced by D1. We could then programmatically draw a subgraph that includes E1 and all its children except for J1 and the leaves as a “module” to be expanded into a single function.

![Screenshot from 2026-02-03 17-33-00.png](work-plan-files/Screenshot_from_2026-02-03_17-33-00.png)

Here’s what the exported Python code for E1, F1, and G1 looks like as a module, as opposed to three separate functions:

```python
# ============================================
# Approach 1: Each cell as a separate function
# ============================================

def F1() -> float:
    return H1() * 2

def G1() -> float:
    return I1() + 5

def E1() -> float:
    return F1() + G1() + J1()

# ============================================
# Approach 2: Single consolidated function
# ============================================

def E1() -> float:
    F1 = H1() * 2
    G1 = I1() + 5
    return F1 + G1 + J1()
```

The consolidated version could even be simplified to a one-liner if we want!

### 3. Deduplication

Assuming you can complete item #2 and detect likely modules, you may also be able to detect subgraphs that result in identical Python code output (except for cell addresses). Then we would only need to output one single function that takes the cell addresses as arguments, and we could reuse that same function in both locations. I don’t know if this is likely to work or not, but it’s worth a try.

### 4. Vectorization

I think there will be subgraphs detected in #2 that operate on themselves in a way that really lends itself to representation as a numpy array. And there might be performance, conciseness, and interpretability gains from representing it that way.

There’s probably also a version of the graph segmentation workflow described above where instead of a single-***cell*** entrypoint to the subgraph, we look for a single-***row*** entrypoint to the subgraph.

But honestly, I haven’t completely thought through how to do this yet. I think I might need to make headway on the easier parts of the problem before I can conceptualize how to make headway on vectorizing the outputs.

**Concrete deliverables:** We may want to be a bit selective about how much time we spend here, depending on time constraints, but let’s assume PRs for items 1 and 2.

**Timeframe:** 4-7 days

# Future phases: Extending library functionality

Obviously Nature Finance has some specific workflows like running alternative scenarios and running sensitivity tests that they want to use this library for. I have mostly been thinking of as library *use cases* rather than library *features*, but I’m sure we could design helpers for these use cases to support them more explicitly and directly. We could also do some demos or tutorials like the ones that often ship with R libraries.

Also, to support Nature Finance’s objective to understand the internal logic of the LIC DSF and how data flows through it, we could write up some analyses of how particular inputs are connected to outputs through the graph. Especially if we manage to programmatically detect and extract some meaningful abstractions during graph segmentation and/or we annotate our functions with informative docstrings using AI, then tracing data through the Python code could end up being quite informative.