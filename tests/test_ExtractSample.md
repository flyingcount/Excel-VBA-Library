# Test: ExtractSample

## Setup

1. Load `ExcelVbaLib.xlam`.
2. On a sheet, put several rows of values (e.g. `A1:C20`).

## Steps

Menu **Excel VBA Lib → Sampling → Extract sample of rows**, or Immediate window:

```vb
Application.Run "ExtractSample", Range("A1:C20"), 50, False
Application.Run "ExtractSample", Range("A1:C20"), 50, True
```

With no arguments, the macro prompts for range, replacement, and percent.

## Expected

- Sheet **Sample** with a title in A1 (count, percent, with/without replacement).
- Sample rows starting at A3. Without replacement, row count is `Fix(percent/100 * n)` (at least 1) and no source row repeats.
- With replacement, source rows may repeat; **Times sampled** can be > 1.
- Draw-order source row numbers, then a frequency list (row number / times sampled).
- Original Personal bug (one fewer row than requested) is fixed.
