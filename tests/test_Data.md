# Test: Data fills (Personal Menu6)

## Setup

1. Load `ExcelVbaLib.xlam`.
2. Do not import `modApiData` / `modInternalData` into the test workbook.
3. Use **Excel VBA Lib → Data** (and **Data → Probability distributions**). Add-in macros do not appear in Alt+F8.
4. If Poisson still raises 438, the loaded add-in is stale. With Excel open and the add-in loaded:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1
   ```

   Use `-XlamPath` if the loaded add-in is not `build\ExcelVbaLib.xlam` under the repo. Then retry Poisson (no Excel restart required).

## Combinations

1. In A1:C3 enter `A,B` / `1,2` / `X,Y` (two values per column, third row blank).
2. Select `A1:C3`. Menu **Excel VBA Lib → Data → Combinations**.

Expected: label to the right of the selection, `Number of combinations = 8`, then the 8 cartesian rows. Cancel of the large-fill prompt (if shown) leaves the sheet unchanged.

## Random fills (select a range first)

Select `E1:E20`, then each of:

| Menu | Prompts | Expected |
|------|---------|----------|
| Random integers | min 0, max 10 | Integers in 0-10 |
| Random numbers | min -1, max 1 | Reals in that interval |
| Random dates | earliest / latest as `yyyy-mm-dd` | Dates formatted `yyyy-mm-dd` in the interval |
| Random strings | min 4, max 8, type 2 | Lowercase strings of length 4-8 |
| Random TRUE/FALSE | (none) | Boolean True/False |
| Random Yes/No | (none) | `Yes` or `No` |
| Random 1 or 0 | (none) | 1 or 0 |

Cancel on any InputBox must leave the range unchanged.

## Random from list

Source list in `G1:G5`. Destination `H1:H20`. Menu **Random from list**: pick destination, then source. Destination cells are values from the source list.

## Yes/No dataset

Menu **Yes/No dataset**: true positives 10, true negatives 8, rows 25.

Expected: sheet **Yes No Dataset**, headers Predicted/Actual, 10 Yes/Yes, 8 No/No, 7 mixed, rows shuffled (not all TP then TN in a block).

## Random test data types

Select `A1:B15`. Menu **Random test data types**. Header row plus Number / Integer / Boolean / Date / String in column A.

## Random Time series

Menu **Excel VBA Lib → Data → Random Time series**.

1. Start-cell InputBox defaults to the active cell. Accept it (or pick another single cell). Cancel leaves the sheet unchanged.
2. Data-points InputBox defaults to 100. Enter `20` (or another whole number ≥ 1). Cancel after choosing a start cell still leaves the sheet unchanged.

Expected, from the start cell:

- Two columns with headers `Date` and `Value`.
- 20 data rows (header + 20 points when the count is 20).
- Dates formatted `yyyy-mm-dd`, consecutive calendar days starting at today.
- Values are a Gaussian random walk (not iid uniform fills). A second run produces a different series.
- A count that would run past the last worksheet row is rejected. Non-integers and counts below 1 are rejected.

## Probability distributions

Select a range (or accept the InputBox range), then **Excel VBA Lib → Data → Probability distributions**:

| Item | Typical params | Expected |
|------|----------------|----------|
| Binomial | trials 10, p 0.5 | Integers 0-10 |
| Bernoulli | p 0.8 | 0 or 1; about 80% ones |
| Normal | mean 0, sd 1 | Reals, roughly bell-shaped |
| Poisson | lambda 4 | Non-negative integers clustered near 4; must not raise 438 |
| Exponential | lambda 1 | Positive reals |
| Gamma | alpha 2, beta 2 | Positive reals |
| Hypergeometric | n 10, K 20, N 50 | Integers 0-10 (successes in the sample) |

Hypergeometric is sequential draws without replacement (not a PDF evaluated at a random k). Cancel on a prompt must not fill.

## Immediate window (optional)

```vb
Application.Run "RandomIntegers"
Application.Run "DataCombinations"
Application.Run "CreateYesNoDataset"
```

These still prompt; there is no argument list (same as Personal).

## Data Preprocessing (Personal Menu5 scaling / dummies)

Use **Excel VBA Lib → Data Preprocessing** (immediately after **Data**, not nested under it). Select numeric columns with **no header**.

### Standardise / normalise / robust

On a sheet starting at A2 (so a title can sit on row 1):

```text
1  10
3  10
5  10
```

A second prompt asks for the output cell. Default is two cells to the right of the last source column (`D2` for `A2:B4`). Accept the default, or pick another cell. Cancel on either prompt leaves the sheet unchanged.

**Standardise columns**: left column ≈ −1.2247, 0, 1.2247 (population SD). Right column all 0 (zero SD). Title **Scaled data (standardised)** on row 1.

**Normalise columns**: left column 0, 0.5, 1. Right column all 0 (zero range). The first column must stay filled (Personal `ReDim` inside the column loop wiped earlier columns).

**Robust scale columns**: (x − median) / IQR. For the left column median 3, IQR = 4, values −0.5, 0, 0.5. Zero IQR → 0.

Blanks and one-row ranges are rejected.

### Convert categorical data to dummy variable matrix

Single column `A, B, A` in A2:A4. **Convert categorical data to dummy variable matrix**. Yes to headers.

Expected: header row on row 1 (`Value`, `A`, `B`); original values in C2:C4; dummy 1/0 in D2:E4 aligned with the source rows. No to headers writes only the dummy matrix at C2, aligned with A2.

A blank cell is skipped as a category and that row is all zeros. Two columns are rejected.
