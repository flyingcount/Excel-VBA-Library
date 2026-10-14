# Infrastructure catalog (Personal.xlsb)

Scan of `_export_raw` (1,308 procedures). Goal: separate **infrastructure** (shared plumbing) from **feature** code (Benford, XmR, menus) and **`z*` templates** (snippet libraries).

## How this was classified

A routine is treated as infrastructure if **any** of these hold:

1. **High call fan-in** — referenced from many modules (e.g. `WriteArrayToWorksheet` ~127 call sites / 41 modules).
2. **High copy fan-out** — same name pasted into many modules (often with **variants**).
3. **Lives in an `aPublic*` hub** — intentional shared library (incomplete, but authoritative intent).
4. **Generic Excel plumbing** — app speed, sheets, named ranges, arrays, range validation, folders, status bar — not domain math.

Excluded from “infrastructure” even if duplicated: feature-specific plots (`BenfordPlot`, `FormatOutputSheet` for a given analysis), forecast stats, cipher UI, etc.

---

## Tier A — Core infrastructure (centralize first)

These are the real shared platform. Prefer **one canonical body** in `Internal`/`Api`, then delete copies.

| Capability | Canonical candidates | Dup modules | Body variants | Approx. call refs / modules |
|------------|---------------------|------------:|--------------:|-------------------------------|
| Write array → sheet | **`aPublicProcedures.WriteArrayToWorksheet`** (newest, typed, safer); also `WriteArrayToWorksheetA1` | 21 | **12** | ~127 / 41 (+ A1 ~64 / 28) |
| Array rank | `GetArrayDimension` — copies nearly identical; **aPublicProcedures has a longer variant** | 21 | **2** | ~43 / 21 |
| App speed (legacy pair) | `TurnOff/OnScreeupdatesAndCalculation` | 18 | **4** | ~90–93 / 56–57 |
| App speed (newer pair) | **`SpeedOn` / `SpeedOff`** in `aPublicProcedures` (saves calc mode) | — | 1 | ~17–28 / 13–14 |
| Output sheet lifecycle | `CreateOutputSheet`, `CheckExistenceAndDeleteOutputSheet`, `DeleteOutputSheet`, `CheckOutputSheetExistenceAndCreate` | 7 | 2 | ~55–59 / 42–44 |
| Sheet exists | `WorksheetExists`, `CheckSheetExists` | 12 | 2 | ~27 / 14 |
| Named ranges | `CheckNamedRangeExists`, `CreateNamedRange`, `*ForSheet` variants | 10–19 | 1–3 | ~22–41 / 10–19 |
| Prepare blank output sheet | `PrepareOutputSheet` | 11 | 3 | ~14 / 11 |

### Important variant notes

- **`WriteArrayToWorksheet` is not one function.**  
  - `aPublicProcedures`: sheet-name + start row/col, existence check, 0/1-base handling, 1D transpose.  
  - `zStockCodeArrays` / many menus: older signature (often address/`ActiveSheet`), thinner error handling.  
  - Plot modules share a common mid-era clone (8 identical bodies).  
  - Histogram / Combinations / Diebold / Cipher each diverge further.  
  → Promote **aPublicProcedures** as v2 API; keep a compatibility wrapper if old signatures are still needed.

- **`TurnOffScreeupdates…` vs `SpeedOn`:**  
  Legacy toggles force Automatic on restore; `SpeedOn`/`SpeedOff` restore prior calc mode. Prefer Speed* as the long-term Internal API; treat TurnOff/On as deprecated aliases.

- **`FormatOutputSheet` (24 copies, 24 variants):**  
  Name sounds shared, but bodies are **feature-specific formatting**. Treat as **feature**, not infrastructure — only extract tiny shared bits (autofit, header bold) if they truly match.

---

## Tier B — Declared public library (`aPublic*`) not fully adopted

These modules **are** infrastructure by design, even when call fan-in is lower (many features never refactored to use them).

| Module | Role | Promote to |
|--------|------|------------|
| `aPublicProcedures` | Sheets, arrays I/O, speed, folders, status, fill-down, blank rows | `modApi*` + `modInternal*` |
| `aPublic_RangeValid` | Numeric/empty/shape checks (`IsRangeNumeric`, `RangesSameNumberOfRows`, …) | `modInternalRangeValid` |
| `aPublic_RangeFn` | Range slice helpers (`ExtractBody`, `RangeColumn`, corner addresses, …) | `modInternalRangeFn` |
| `aPublic_ArrayValid` | Row/column vector tests, transpose helper | `modInternalArrayValid` |
| `aPublic_ArrayFn` | Covar/corr/SD vectors, reshape — **numeric array utils** (infra for stats features) | `modInternalArrayFn` or keep as UDF-adjacent |
| `aPublic_InputBox` | `GetRange` / `GetLong` / … input prompts | `modApiUi` / `modInternalInput` |
| `aPublic_Tables` | ListObject create/clear/exists/list | `modApiTables` |
| `aPublicFunctions` | Misc (`sprintf`, colour index, strip chars, worksheet name gen, …) | split: text → InternalText; sheet name → ApiSheets |
| `aPublicDialogBoxes` | `SelectFolder` | `modApiFiles` |
| `aPublic_Format` | `ColumnsAutoFit` | `modInternalFormat` |

**Gap:** Tier A call sites often use **private clones** instead of these modules — so `aPublic*` is the intended hub but not the complete runtime graph.

---

## Tier C — Parallel infra hubs (dedupe against aPublic*)

| Module | Overlap | Action |
|--------|---------|--------|
| `NamedRanges` / `Fn_NamedRanges` | Same ideas as Create/Check named range in menus + aPublic | Merge into `modInternalNamedRanges`; keep `Fn_*` only if used as worksheet UDFs |
| `Filehandling` | `SelectFolder2`, `Open_Workbook`, `CreateFolder` | Merge with `aPublicDialogBoxes` / `aPublicProcedures` folder APIs |
| `ResizeRange` | Range geometry utilities | Review vs `aPublic_RangeFn` |
| `vbaCreateSheet` / `vbaClearSheet` (in Symbols / Fn_Gini) | Alternate sheet helpers | Alias to ApiSheets |

---

## Tier D — `z*` modules = templates / stock snippets (not production infra)

| Module | Contents | Treat as |
|--------|----------|----------|
| `zStockCodeArrays` | `WriteArrayToWorksheet`, `GetArrayDimension`, `GetArraySlice`, `UploadTableToArray`, extract rows/cols | **Template** of array I/O — harvest best bits into Internal; do not ship module as-is |
| `zStockCode` | Named-range + `CaptureInputRange` snippets | Template for named-range Internal |
| `zStockCodeRangeVal` | `IsRangeValid`, numeric/square checks | Template overlapping `aPublic_RangeValid` |
| `zStockCodeCharts` | Quick chart snippets | Feature template |
| `zStockCodeMsgBox` | MessageBox sample | Template |
| `zCollectionTools` / `Z_Collections` | Collection ↔ range/sheet | Optional Internal Collections (low fan-in today) |
| `z_WIP` | Experiments + a local `FormatOutputSheet` | Exclude from library until promoted deliberately |

Use `z*` as **design references** when reconciling variants; features should call curated Api/Internal, not `z*`.

---

## Tier E — Not infrastructure (despite duplication)

| Pattern | Why |
|---------|-----|
| `FormatOutputSheet`, `PrepareOutputSheet` (when formatting analysis output) | Per-feature layout |
| `BenfordPlot` / `CreateBenfordTable` / `BenfordAnalysisPlot` | Domain feature family |
| `PlotCorrelogram*`, residual/QQ plot formatters | Stats feature |
| `MatrixMultDefined` across matrix menus | Domain; may share a **math** library later, separate from sheet I/O |
| Trig duplicates in `Fn_Flying` vs `Fn_Trigonometry` | UDF domain overlap, not Excel plumbing |

---

## Suggested infrastructure surface (target library)

```text
Internal/
  modInternalExcelApp      ← SpeedOn/Off (+ deprecated TurnOff/On aliases)
  modInternalSheetIO       ← WriteArrayToWorksheet*, WriteRangeToWorksheet, GetArrayDimension
  modInternalOutputSheets  ← Create/Delete/Check output sheet
  modInternalNamedRanges   ← Check/Create named ranges (* / *ForSheet)
  modInternalRangeValid    ← from aPublic_RangeValid (+ zStockCodeRangeVal ideas)
  modInternalRangeFn       ← from aPublic_RangeFn
  modInternalArrayFn       ← dimension/slice/extract (+ zStockCodeArrays Upload/Extract)
  modInternalText          ← Remove_Numbers, sprintf, ValidRangeName helpers
  modInternalFormat        ← ColumnsAutoFit, shared header formatting only

Api/
  modApiArrays             ← Public wrappers over SheetIO
  modApiSheets             ← WorksheetExists, UserSelectWorksheet, output sheet API
  modApiFiles              ← SelectFolder, CreateFolder/DeleteFolder
  modApiTables             ← aPublic_Tables
  modApiUi                 ← StatusUpdate/Clear, InputBox helpers, MsgBox
```

---

## Recommended reconciliation order

1. Freeze **canonical** bodies from `aPublicProcedures` for Speed* + WriteArray* + output sheets.  
2. Diff the **12 WriteArray variants**; document signature matrix (params / ActiveSheet vs named sheet / A1 helper).  
3. Implement compatibility shims only where pilots still need old signatures.  
4. Merge range validation: `aPublic_RangeValid` ∪ unique bits from `zStockCodeRangeVal`.  
5. Replace private clones in 2–3 pilot menus; measure compile + behaviour.  
6. Leave `FormatOutputSheet` alone until shared formatting lines are proven identical.

---

## Artefacts from the scan

| File | Purpose |
|------|---------|
| `_export_raw/inventory.csv` | Procedure index |
| `_export_raw/proc_bodies.csv` | Per-proc body hash / length (variant detection) |

Re-run the scan after large Personal edits. Update this doc when a canonical variant is chosen (record the winning `BodyHash` / source module).
