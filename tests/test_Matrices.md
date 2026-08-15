# Test: Matrices (Personal Menu13)

Personal dump (`Custom_Menu13_*`, `Fn_Matrices*`) is gitignored and was not in the cloud workspace. This pack covers the same families (create, operations, Cholesky, symmetric eigen, QR) with array writes and cancel-safe prompts.

## Setup

1. Load `ExcelVbaLib.xlam`.
2. Do not import matrix modules into the test workbook.
3. Use **Excel VBA Lib → Matrices**. Add-in macros do not appear in Alt+F8.

## Create

Select an empty cell. **Matrices → Create → Identity**, rows = 3.

Expected: 3 x 3 identity starting at that cell. Zeros / Ones / Hilbert / Random follow the same pattern. **Diagonal from vector**: select a column of 3 numbers; a 3 x 3 diagonal appears to the right.

## Operations

Put `= {1,2;3,4}` equivalent in A1:B2 (1 2 / 3 4). Select A1:B2.

| Menu | Expected to the right of the selection |
|------|----------------------------------------|
| Transpose | 1 3 / 2 4 |
| Inverse | -2 1 / 1.5 -0.5 |
| Determinant | -2 |
| Trace | 5 |
| Frobenius norm | sqrt(1+4+9+16)=sqrt(30) |
| Is symmetric | FALSE |

**Multiply**: A1:B2 times itself, second matrix = A1:B2. Result 7 10 / 15 22.

Cancel on the second-matrix InputBox must not write.

## Decompositions

- **Cholesky** on a SPD matrix such as 4 2 / 2 3. Lower L to the right; `L * L^T` recovers A.
- **Eigen (symmetric)** on 2 1 / 1 2. Last column eigenvalues near 3 and 1; A v ≈ λ v.
- **QR** on a tall or square numeric range. Q on top, R immediately below; Q^T Q ≈ I and Q R ≈ A.

## Worksheet UDFs

Array-enter or use dynamic arrays:

```
=MatMult(A1:B2,A1:B2)
=MatInv(A1:B2)
=MatDet(A1:B2)
=MatChol(A1:B2)
=MatrixMultDefined(A1:B2,A1:B2)
```

Empty or text cells → `#VALUE!`. Non-symmetric Cholesky / Eigen → `#VALUE!` (UDF) or an error message (menu).

## Limits

Side length above 250 is rejected. Results write one column to the right of the selection (create writes at the active cell).
