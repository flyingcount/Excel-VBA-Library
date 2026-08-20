# source/

These `.bas` files are a snapshot of what was imported into **`build/ExcelVbaLib.xlam`**.

**The add-in is the library.** Add new modules in the VBE of `ExcelVbaLib.xlam` (`Alt+F11` → Insert → Module). Do not create new git folders here for Personal123 families.

`_export_raw/` is a local dump from `Data/Personal123.xlsb` (gitignored). It is not imported into the add-in.

To rebuild the current add-in from this snapshot:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
```

To replace Data modules in a loaded add-in (`modInternalData`, then `modApiData`):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1
```

That does **not** update Matrices. Refresh the whole add-in with `-All`:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Import-AddinModules.ps1 -All
```

That writes `modApiMatrices1` and `modApiMatrices2` plus the rest of Menu13 (`modApiMatrixCreate`, `modApiCholesky`, `modApiEigenDecomp`, `modApiUnitary`, `modApiMatrixUtilities`). Excel VBA Lib → Matrices holds Create / Operations / Properties / Validation / Decompositions (Personal-only tools mixed in; duplicate Personal identity/zeros/ones/exchange/transpose/dot entries are omitted). **Analysis** (Personal Menu18 plus Menu5 Bland-Altman / Deming / Box-Cox / logit) is SVD, AX=B, confusion, covariance/correlation, residuals, and those method-comparison tools. **Data → Data Preprocessing** (Menu5) is column standardise / normalise / robust scale and dummy variables. **Plots Charts** (Personal Menu11) is histograms, regression, capability, distribution plots, QQ, Lorenz/Gini, ACF, Diebold-Mariano, and XmR. **Ranges** (Personal Menu14) is named-range inventory/create/hide/scope, unique values, analysis, frequency, cleanse, transpose, and split-to-named-columns. **Duplicates** (Personal Menu15) is flag / letter-label / colour / count duplicate cells. **Protection** (Personal Menu7) is scroll-area limits, default-password protect/unprotect, and unhide. **Reconciliations** (Personal Menu20) is two-column number/string recon and range compare. **Time series** (Personal Menu27) is ACF/PACF analysis and lag differencing.
