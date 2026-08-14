# Test: `CreateDateTable`

## Setup

1. Load `ExcelVbaLib.xlam` — see [build/README.md](../build/README.md).
2. Do not import date-table modules into the test workbook.

## Steps

Immediate window (add-in loaded):

```vb
Application.Run "CreateDateTable", #1/1/2024#, #1/31/2024#
Application.Run "CreateDateTable", #4/1/2024#, #3/31/2025#, "DimDate", 4
```

## Expected (first call)

- Sheet `DateTable` exists with header row + 31 data rows.
- Columns: Date, Year, Quarter, Month Number, Month, Day, Day Of Week Number, Day Of Week, Week, MMM-YYYY, Month sort, Fiscal Year, Fiscal Period.
- `Date` formatted `yyyy-mm-dd`; Monday = 1 for Day Of Week Number.
- Excel Table present (default `AsExcelTable:=True`).

## Expected (second call, FY start April)

- Sheet `DimDate`.
- Fiscal Year / Fiscal Period reflect April start (e.g. 2024-04-01 → Fiscal Period 1).
