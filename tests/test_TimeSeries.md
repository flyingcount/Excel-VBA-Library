# Test: Time series (Personal Menu27)

Personal `Custom_Menu27_TimeSeries` / `Custom_Menu27_TS_DateDiff`, rewritten as `modApiTimeSeries` + `modInternalTimeSeries`.

Private Personal tools `GeneratePACFList` / `GenerateAVCFList` / `GenerateAutoCovarianceMatrix` are not on the menu.

**Plots Charts → Autocorrelation (ACF)** is a separate correlogram (`GenerateCorrelogram`). This pack is the Menu27 analysis sheet plus lag differencing.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Time series**. Do not import Personal `Custom_Menu27_*` into the test workbook (Personal also defines `ACF` / `PACF` UDFs).
3. Work in a throwaway workbook. Each tool replaces its output sheet.

Cancel on the range InputBox does nothing. Entire-column selections use UsedRange. Personal turned screen updating off **before** the InputBox, so Cancel could leave Excel frozen. Non-numeric or blank cells are rejected. A two-column selection is rejected (Personal analysis allowed it, then `ACF` returned an error string).

## Time series analysis

```
A1: 10
A2: 12
A3: 11
A4: 13
A5: 12
A6: 14
A7: 13
A8: 15
A9: 14
A10: 16
```

Select `A1:A10`. **Time series → Time series analysis**.

Expected: sheet **Time Series**.

- A1 title **Time series analysis**. Data in A6:A15. Sheet-scoped name `Time_Series_Data`.
- Headers at A5:M5 (Lag, ACF, ACVF, PACF, Bartlett, Box-Pierce, Ljung-Box, …).
- Lags 0..2 (`n\3` for n=10). Lag 0 has ACF ≈ 1 and ACVF (variance); PACF and the portmanteau tests are blank.
- Lag 1+ has PACF, Bartlett (`|ACF|` vs the critical value; Personal tested only ACF > crit), Box-Pierce and Ljung-Box.
- Box-Pierce / Ljung-Box **reject white noise** when p < 0.05. Personal said “statistically equal to zero” in that case (the hypothesis was inverted).
- Duke / Real Statistics links in G1:G4 and J1.

Need at least 3 numeric rows. Personal used unqualified `Cells` after creating the sheet.

## Time series analysis with formulae

Same data. **Time series → Time series analysis with formulae**.

Expected: same sheet layout, lags 0..1 (`n\4` for n=10). ACF/PACF/… cells are A1 formulas that call this add-in’s UDFs (`='ExcelVbaLib.xlam'!ACVF(Time_Series_Data,C6)`, and so on). Excel 365 may display `=@ACVF(...)`. Unqualified `=ACVF(...)` works after the add-in menu has registered the UDFs (Insert Function category **Excel VBA Lib**); before that it is `#NAME?`. Changing A6 should recalculate.

## Date differencing

Same column. **Time series → Date differencing**.

Expected: sheet **Date Diff**. This is lag-k differencing `y[t]-y[t-k]` of the original series (pandas `Series.diff(lag)`), not calendar date subtraction and not repeated first differences (Δ², Δ³).

- Slope / Intercept / R² in rows 21–23.
- Table from row 25: Index, Time series data, then lags 1..`min(Round(n/3),24)`.
- Leading rows that cannot be differenced are blank. Lag 1 has **9** values for n=10 (rows 2–10). Lag 2 is `y[t]-y[t-2]` (eight 1s on this series), not Δ².
- Charts at A4 / B4 / L4 (original, lag-1, lag-2) with a red trendline. No cell is selected to place the chart.
- Link in D2 to Real Statistics ARIMA differencing.
- Below the data, an **Autocorrelation** table (values, not formulas) has lag 0..n/4 for the original series and every difference column (used cells only, skipping leading blanks). Lag-1 ACF is green when ≥ −0.5 and red when < −0.5. A flag row says `OK` or `Over-differenced`.

## Date differencing with formulae

**Time series → Date differencing with formulae**.

Expected: sheet **Date Diff Formula**. Difference cells are R1C1 formulas of the original series: lag k is `=RC[-k]-R[-k]C[-k]` (lag 1 is `=RC[-1]-R[-1]C[-1]`). Changing the original series updates the diffs. Same values as Date differencing / `difference(data, lag=k, order=1)`.

Below the data, an **Autocorrelation** table has lag 0..n/4 for the original series and every difference column (`='ExcelVbaLib.xlam'!ACF(...)` on the used cells only). Lag-1 ACF cells are green when ≥ −0.5 and red when < −0.5 (over-differenced). A flag row under the table says `OK` or `Over-differenced`.

A single-cell column is allowed for differencing (n≥2 required).
