# Test: Benford analyses

## Setup (preferred)

1. Build and load `ExcelVbaLib.xlam` — see [build/README.md](../build/README.md).
2. Do not import Benford modules into the test workbook.

## Setup (modules only, if not using the add-in)

1. Import, in order:
   - `source/Internal/modInternalExcelApp.bas`
   - `source/Internal/modInternalError.bas`
   - `source/Internal/modInternalNamedRanges.bas`
   - `source/Internal/modInternalBenford.bas`
   - `source/Api/modApiBenford.bas`
2. Compile (`Debug → Compile VBAProject`).

## Sample data

On a sheet, put mixed numeric values in `A2:A50` (include some non-numeric cells, zeros, and decimals).

## Steps

Immediate window (add-in loaded):

```vb
Application.Run "BenfordAnalysisFirstDigit", Range("A2:A50")
Application.Run "BenfordAnalysisSecondDigit", Range("A2:A50")
Application.Run "BenfordAnalysisThirdDigit", Range("A2:A50")
Application.Run "BenfordAnalysisTwoDigit", Range("A2:A50")
Application.Run "BenfordAnalysisThreeDigit", Range("A2:A50")
Application.Run "BenfordAnalysisLastTwoDigit", Range("A2:A50")
```

Calling any of these with no argument should show the range InputBox (cancel must restore screen updating / calculation).

Worksheet formulas:

```excel
=Benford2ndDigitProbability(0)
=Benford3rdDigitProbability(5)
```

## Expected

| Call | Sheet created | Digit column | Charts |
|------|---------------|--------------|--------|
| First | Bedford Analysis First digit | First digit (1–9) | Frequency, actual vs Benford, residuals |
| Second | Bedford Analysis Second digit | Second digit (0–9) | Frequency, actual vs Benford |
| Third | Bedford Analysis Third digit | Third digit (0–9) | Frequency, actual vs Benford |
| Two-digit | Bedford Analysis 2 digit | First two digits (10–99) | Frequency, actual vs Benford |
| Three-digit | Bedford Analysis 3 digit | First three digits (100–999) | Frequency, actual vs Benford |
| Last two | Bedford Analysis last 2 digits | Last two digits (00–99), expected frequency 0.01 | Frequency, actual vs Benford |

- Summary table includes Count / Actual frequency / Benford (or expected) frequency / Z-statistic.
- Z-statistics above the confidence Z-score (`$J$2`) are highlighted.
- `Benford2ndDigitProbability(0)` is about 0.1197; probabilities for digits 0–9 sum to 1.
