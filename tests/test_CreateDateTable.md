# Test: `CreateDateTable`

## Setup

1. Import `modInternalDateTable`, `modInternalSheetIO`, `modInternalExcelApp`, `modInternalError`, `modApiSheets`, `modApiDates` into `ExcelVbaLib.xlam` (or a test workbook).
2. Compile (`Debug → Compile VBAProject`).

## Steps

In the Immediate window (or a caller workbook with the add-in loaded):

```vb
CreateDateTable #1/1/2024#, #1/31/2024#
CreateDateTable #4/1/2024#, #3/31/2025#, "DimDate", 4
```

## Expected (first call)

- Sheet `DateTable` exists with header row + 31 data rows.
- Columns: Date, Year, Quarter, Month Number, Month, Day, Day Of Week Number, Day Of Week, Week, MMM-YYYY, Month sort, Fiscal Year, Fiscal Period.
- `Date` formatted `yyyy-mm-dd`; Monday = 1 for Day Of Week Number.
- Excel Table present (default `AsExcelTable:=True`).

## Expected (second call, FY start April)

- Sheet `DimDate`.
- Fiscal Year / Fiscal Period reflect April start (e.g. 2024-04-01 → Fiscal Period 1).
