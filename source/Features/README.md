# Features

Domain tools from Personal `Custom_Menu*` that are not generic plumbing. Land a family here first (call graph intact). Split to `Api/` + `Internal/` later if it becomes a public library surface (as Benford did).

## Subfolders

| Folder | Personal123 modules (typical) |
|--------|-------------------------------|
| [Stats/](Stats/) | `Custom_Menu5_*` (except Benford, already in Api), `Custom_Menu11_*` plots/XmR, `Custom_Menu18_*` analysis |
| [Matrices/](Matrices/) | `Custom_Menu13_*`, matrix helpers that are not worksheet UDFs |
| [Editing/](Editing/) | `Custom_Menu1_*`, `Custom_Menu14_*`, `Custom_Menu15_*`, `Custom_Menu16_*` |
| [Charts/](Charts/) | `Custom_Menu3_*` |
| [Output/](Output/) | `Custom_Menu8_*` |
| [Workbook/](Workbook/) | `Custom_Menu2_*` inventory / properties / hide sheets |
| [TimeSeries/](TimeSeries/) | `Custom_Menu27_*`, `Custom_Menu28_*` forecasts |
| [Other/](Other/) | Cipher, primes, sampling, finance, PQ UI, anything without a folder yet |

Benford is **not** here — it lives in `../Api/modApiBenford.bas` + `../Internal/modInternalBenford.bas`.

While landing, Personal module names (`Custom_Menu11_Histogram.bas`) are fine. Rename to `modHistogram.bas` when the public surface is stable.
