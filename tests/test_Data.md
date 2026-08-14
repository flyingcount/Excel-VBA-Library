# Test: Data fills (Personal Menu6)

## Setup

1. Load `ExcelVbaLib.xlam`.
2. Do not import `modApiData` / `modInternalData` into the test workbook.
3. Use **Excel VBA Lib → Data** (and **Data → Probability distributions**). Add-in macros do not appear in Alt+F8.

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

## Probability distributions

Select a range (or accept the InputBox range), then **Excel VBA Lib → Data → Probability distributions**:

| Item | Typical params | Expected |
|------|----------------|----------|
| Binomial | trials 10, p 0.5 | Integers 0-10 |
| Bernoulli | p 0.8 | 0 or 1; about 80% ones |
| Normal | mean 0, sd 1 | Reals, roughly bell-shaped |
| Poisson | lambda 4 | Non-negative integers |
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
