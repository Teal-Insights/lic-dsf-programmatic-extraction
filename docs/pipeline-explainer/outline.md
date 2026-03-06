Problem:

Excel workbooks are powerful — but the logic runs *inside the application*. To run scenarios, economists must open Excel, change cells, and look at the results.

This is *laborious*, *time-consuming*, and *error-prone*. Workflows cannot be *typed*, *tested*, or even easily *audited*. This can't be *automated* or done *at scale*.

I deally, we would *migrate the workbook's logic to Python*. But that's a big challenge. The LIC DSF template is *complicated*.

The LIC DSF has *200,000 formula cells* that call 55 unique functions, nested up to 7 levels deep. It uses lookups and offsets that are *tightly coupled to Excel's data model*.

Solutions:

Option 1: *drive Excel from Python*. Pros: modest speedup, familiar tooling. Cons: still has all the limitations of Excel, can't be productionized for use in other applications.

Option 2: *have AI agents translate to Python*. Pros: leverages the "bitter lesson," improves as AI models do. Cons: unpredictable output shape, no correctness guarantees, unreproducible for new template versions.

Option 3: *transpile to Python*. Pros: predictable output shape, correctness guarantees, reproducible for new template versions. Cons: hard to get right, output is still Excel-shaped.

AI agents help make Option 3 tractable. And we can *post-process the output to make it more Python-shaped*.

For interpretability, reproducibility, and correctness guarantees, *don't generate the code*. *Generate the code that generates the code*.

