# `py-lic-dsf` library design insights from the LIC DSF course and guidance note

## Insights from the Instructions worksheet

The Instructions worksheet adds several useful design signals beyond the guidance-note/course notes.

- **There are two kinds of inputs**: policy/evaluation settings (e.g., debt definition, coverage, vintage year) and numeric series assumptions. Treat these as separate top-level categories.
- **Workflow is explicitly staged** (`Input 1` → `...` → `Output 7`), so the API should mirror that sequence rather than a flat list of setters.
- **Some sections are conditional/optional** (`Input 6 Standard Test`, `Input 6 Tailored Tests`, customized scenarios). This suggests `required_if(...)` metadata on setters.
- **Judgment is part of the official flow for the external debt risk rating**. So the output we want to target for graph extraction is the *mechanical* rating, not the final indicator. That lets us skip the judgment steps in the library API.
- **Template behavior includes auto-population/defaults**. We should see whether this is done via formulas or macros, and replicate the macros if necessary.
- **Domain constraints are documented in instructions** (e.g., avoid negative residual financing, nominal vs real rate conventions). These should become runtime validation rules.

Rating apparently lives on `Output 7`.

## Insights from the LIC DSF course and guidance note

Design principle: **shape the API around DSF workflow decisions, not Excel sheet layout**.

- **Core abstractions**: baseline assumptions, stress tests, thresholds, and ratings.
- **Time semantics matter**: Split by historical data vs projection assumptions vs shock overrides and other parameter configs; this should be first-class in grouping.
- **External vs public lens**: Document whether variables feed external DSA, public DSA, or both, but not in the top-level namespace.
- **Judgment-aware model**: We are only replicating the mechanical rating, but we may want to document or even mechanically flag where use of judgment is appropriate.

For the specific “which categories should we pick?” question, here is a proposed **v1 category set**:

1. `configuration` (`Input 1`, debt definition, vintage, debt coverage)  
2. `macro_debt_baseline` (`Input 3`)  
3. `baseline_financing_terms` (`Input 4` + `Input 5`)  
4. `stress_tests_standard` (`Input 6 standard`)  
5. `stress_tests_tailored_and_custom` (`Input 6 tailored` + customized sheets)  
6. `stress_financing_and_sdr` (`Input 7` + `Input 8`)

This set tracks how analysts actually prepare a DSA (history -> baseline -> financing/debt structure -> shocks -> judgment), while still being compact enough to navigate.

A practical rule for assigning each setter to a category:

- **Primary key** = “What decision is this input used for?”  
- **Secondary tags** = external/public, stock/flow, history/projection, mandatory/optional.

That gives stable namespaces for usability plus richer metadata for docs/search later.
