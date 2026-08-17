# Test: Analysis (Personal Menu18)

Personal `Custom_Menu18_*` rewritten as `modApiAnalysis` / `modApiCovariance` / `modApiConfusion` / `modApiResiduals` / `modApiSvd` / `modApiLinearSystem` plus Internal helpers.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Analysis**. Do not import Personal `Custom_Menu18_*` into the test workbook.

## Mean / standard deviation vectors

On a sheet, header `A B` in A1:B1 and numeric body:

```text
A  B
1  2
3  4
5  6
```

**Analysis → Mean vector**, output default (below the data).

Expected: title **Mean Vector**, header row, then `3` and `4`.

**Stdev (population) vector** on the same block: `√(8/3)` and `√(8/3)` ≈ 1.633. **Stdev (sample) vector**: `√4` = 2 for both columns.

## Variance-covariance and correlation

Same block. **Variance-covariance** writes a 2×2 **population** matrix (divide by n=3), with headers on both axes.

**Correlation matrix** is population covariance divided by products of population SDs; diagonal is 1.

**Standardised covariance** (also on **Matrices → Operations**) is sample covariance of standardised columns (n-1). A bold **Standardised Covariance** label sits above the matrix. The input does **not** have to be square (Personal `MatrixCovariance` required a square range, which is wrong for observations × variables).

**Prove var-covar and correlation** writes sheet **Cov and Correl**: `D × R × D =` population covariance, with a **Check** copy of the direct covariance and a **Variance** row.

## SVD

Select a numeric matrix (e.g. 3×2). **Analysis → Calculate SVD** (also **Matrices → Decompositions → SVD**).

Expected: **Singular Values** as a row, then **Left matrix U**, then **Right matrix V**. Cancel on either prompt does nothing. Non-numeric cells are rejected. Personal required the *output* cell to be numeric and grey-filled a large block; this version does not.

## Solve AX=B

Square A larger than 1×1, B with the same number of rows. **Analysis → Solve AX=B**.

Expected: X written at the output cell. Singular A raises an error. Personal `LinearSystem_AXB_v2` showed a RefEdit form and duplicated Gauss-Jordan (integer path used an undefined `MCM_`); both `LinearSystem_AXB_v1` and `_v2` now use the add-in `Solve`.

## Confusion matrix

Two columns, Predicted then Actual. No header row.

Yes/No (case-insensitive; `Y`/`N` allowed):

```text
Yes  Yes
Yes  No
No   Yes
No   No
```

**Confusion matrix (Yes/No)** → sheet **Confusion Matrix** with a row listing, TP/FP/FN/TN counts, and metrics. F-beta is computed in VBA (Personal wrote a broken A1 formula that never calculated).

**Confusion matrix (1s and 0s)** is the same with 1 = positive class.

**Confusion matrix template** → sheet **Confusion Matrix Template** with live formulae in the metric cells (Personal dumped formula *text* via `WriteArrayToWorksheet`, so nothing calculated). TP/FP/FN/TN cells are unlocked; the sheet is protected.

## Residuals

One numeric column, at least 3 cells. **Residuals analysis** → sheet **Residuals Analysis**: Order 1..n, residuals, sum, SLOPE/INTERCEPT/RSQ, scatter with size-3 **x** markers and a trendline. Named ranges `Order` and `Residuals` are sheet-scoped (Personal used workbook names and unqualified `Range`, so writes could land on the wrong sheet). Blanks are rejected (`IsNumeric` treats empty as numeric).
