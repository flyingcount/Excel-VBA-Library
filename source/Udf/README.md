# Udf

Worksheet-callable functions. **Keep Personal `Fn_*` Public Function names** so existing workbook formulas do not break.

| Personal module | Notes |
|-----------------|--------|
| `Fn_Ageing`, `Fn_TaxUK`, `Fn_UKVAT` | domain UDFs |
| `Fn_Statistics`, `Fn_Zscore`, `Fn_Skew`, `Fn_Gini*` | stats UDFs |
| `Fn_Matrices*`, `Fn_Flying*`, `Fn_Dates2` | keep names; shared math may later move to Internal |
| `Fn_NamedRanges` | prefer `modInternalNamedRanges` for macros; keep this file only if cells call it |

File name = `Fn_{Area}.bas` with matching `Attribute VB_Name`.
