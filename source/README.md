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

That writes `modApiMatrices1` and `modApiMatrices2` plus the rest of Menu13 (`modApiMatrixCreate`, `modApiCholesky`, `modApiEigenDecomp`, `modApiUnitary`, `modApiMatrixUtilities`). Excel VBA Lib → Matrices holds Create / Operations / Properties / Validation / Decompositions (Personal-only tools mixed in; duplicate Personal identity/zeros/ones/exchange/transpose/dot entries are omitted).
