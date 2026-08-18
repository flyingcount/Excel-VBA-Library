# Test: Plots Charts (Personal Menu11)

Personal `Custom_Menu11_*` rewritten as `modApiHistogram`, `modApiDistPlots`, `modApiQQPlots`, `modApiLinearRegression`, `modApiLorenz`, `modApiAcf`, `modApiDiebold`, `modApiXmR`, `modApiProcessCapability`, `modApiChartSheet`, plus `modInternalPlots`.

## Setup

1. Load `build/ExcelVbaLib.xlam`. Fully quit Excel and restart if the menu was just rebuilt.
2. Use **Excel VBA Lib → Plots Charts**. Do not import Personal `Custom_Menu11_*` into the test workbook.

Cancel on any prompt must leave the workbook unchanged (screen updating restored).

## Histogram

Numeric column, e.g. 20 values in A1:A20. **Histogram and data table**, 10 bins.

Expected: sheet **Histogram**, summary in A3:B9, table Index/From/To/Frequency, column chart **Frequency distribution**. Check total is OK.

**Histogram with formulae** → sheet **Histogram formulae**. Bin start / bin range cells are labelled `<- change`. Frequencies are COUNTIF formulae, so they update if `Histogram_Data` changes.

## Linear regression

Two columns x,y with a header-free numeric body, at least 3 rows:

```text
1  2
2  4.1
3  5.9
4  8.2
```

**Linear regression** → sheet **Linear regression**: stats at A5 (confidence 0.95, yellow), interactive x at B16 (yellow), forecast increment B22=10, actuals at M3, 50 forecast rows, chart with actuals + prediction bands + forecast.

**Linear regression v2** → sheet **Linear regression V2**: original data at Y3, integer-x fitted grid at Q3, 100-point forecast (increment 1), residuals at AB, slope CI at N1, residuals chart at D21. Original-data named range is `n×2` (Personal sized it `n×n`).

## Process capability

Numeric process data. **Process capability**, LSL −10, USL 10.

Expected: sheet **Process Capability Plot**, Cp/Cr/Cpu/Cpl/Cpk as live formulae (no `personal.xlsb` UDFs). If you enter LSL > USL they are swapped. Chart overlays the normal curve, spec limits, and normalised frequency.

## Parametric plots

Select a cell on row 2 or below (row 1 is reserved for headers). Use **Plots Charts → Parametric data and plots** (logistic is here, not on the top-level Plots Charts menu).

| Menu | Default inputs (yellow) | Charts |
|------|-------------------------|--------|
| Binomial | p=0.8, 100 trials | PMF + CDF |
| Normal | mean 0, sd 0.5 | PDF + CDF |
| Log-normal | new sheet **Log Normal** | PDF + CDF |
| Poisson | mean 10 | PMF + CDF |
| Weibull / Gamma | alpha 1, beta 10 | PDF + CDF |
| Beta | alpha=beta=0.5, dx=0.01 | PDF + CDF |
| Exponential | sheet **Exponential**, lambda 1.5 | PDF + CDF |
| Hypergeometric | sheet **Hypergeometric**, n=100, K=50, N=500 | PMF + CDF |
| Logistic curve | sheet **Logistics curve** | one sigmoid chart |

Change a yellow cell and the formulae / charts update. R1C1 formulae are written with `FormulaR1C1` (Personal dumped them via `.Value`).

## QQ plots

At least 3 numeric cells. **Gaussian / normal** → sheet **QQ Normal chart** (sorted data, NORMSINV ranks, QQ + residuals charts, slope/intercept/R² as formulae). **Uniform** → sheet **QQ Uniform chart** (theoretical grid from min to max).

Blanks are rejected (`Count` vs cell count; Personal `IsNumeric` treated empty as numeric).

## Lorenz / Gini

Single non-negative numeric column. **Lorenz curve and Gini** → sheet **Gini**, Gini coefficient in J6, Lorenz vs 45° equality line.

Negatives or extra columns are rejected.

## Autocorrelation

Single numeric column, at least 4 rows. **Autocorrelation (ACF)** → sheet **Correlogram**: time series, ACF with Bartlett bands, PACF (Yule-Walker, capped at 40 lags). Alpha is yellow at AD7.

Personal labelled a crude “all values ≥ first observation” flag as stationarity; this version labels it **Non-decreasing from first obs**.

## Diebold-Mariano

- **1 column:** actuals; forecast 1/2 are RAND() placeholders (note on the sheet, yellow). Replace them or re-run with three columns.
- **3 columns:** actual, forecast 1, forecast 2.

Sheet **DieboldMariano**. DM p-value is two-sided (`NORMSDIST` of the absolute statistic). HLN comment uses p-value vs Alpha (Personal compared the HLN statistic to the p-value).

## XmR

Two columns, header row + at least 2 numeric values in column 2. **XmR chart** → sheet **XmR**: X̄ ± 2.66 MR̄ and MR UCL = 3.267 MR̄ as named-range formulae (same constants as Personal, without the relative-average offset maze).

**XmR diagnostics** → sheet **XMR Diagnostics** table `tbl_XmR_Analysis` (31 columns, Western Electric rules, X and mR charts). One or two columns, no header. X̄ / MR̄ / SD average the data block (`$B$4:$B$last`), not sheet row 2.

## Line chart sheet

One column with a header. **Line chart on a chart sheet** → chart sheet named from the header, SE bars and a trendline. Re-running deletes the previous chart sheet of that name.
