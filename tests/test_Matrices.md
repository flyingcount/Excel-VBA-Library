# Test: Matrices (Personal Menu13)

Personal dump (`Custom_Menu13_*`, `Fn_Matrices*`) is gitignored and was not in the cloud workspace. This pack covers create, operations, Cholesky, symmetric eigen, QR, and LU with array writes and cancel-safe prompts.

## Setup

1. After `git pull`, `build/ExcelVbaLib.xlam` is in the repo. To rebuild from source **in Windows PowerShell** (not the cloud Linux terminal):

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-AddinModules.ps1 -All
   ```

The script should list `modApiMatrices1` / `modApiMatrices2` public Subs and fail if those modules are absent.

   Then overlay the original Personal modules (so the VBE shows the Personal Matrices1/2 code, not only the git snapshot):

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Import-Menu13FromPersonal.ps1
   ```

   Personal.xlsb must be in `Data\` (or pass `-WorkbookPath`). Use `-AllMenu13` for the rest of the family.

2. Load `ExcelVbaLib.xlam`. Fully quit Excel and restart so `RegisterMatrixUdfs` runs.
3. Do not import matrix modules into the test workbook.
4. Use **Excel VBA Lib → Matrices**. Add-in macros do not appear in Alt+F8.

## Create

Select an empty cell. **Matrices → Create → Identity**, rows = 3.

Expected: 3 x 3 identity starting at that cell. Zeros / Ones / Hilbert / Random / Exchange follow the same pattern.

- **Diagonal from vector**: select a column of 3 numbers; a 3 x 3 diagonal appears to the right.
- **Toeplitz from vector**: select `3; 1; 0` as a column. Expected symmetric Toeplitz with 3s on the diagonal.
- **Vandermonde from vector**: select `1; 2; 3`, columns = 3. First column is ones.
- **Companion from vector**: select `1; -2; 3` (constant term first). 3 x 3 companion with 1s on the subdiagonal.
- **Vec**: A1:B2 → one column 1 / 3 / 2 / 4 (column-major). **Unvec** with rows = 2 recovers A.

## Operations

Put 1 2 / 3 4 in A1:B2. Select A1:B2.

| Menu | Expected to the right of the selection |
|------|----------------------------------------|
| Transpose | 1 3 / 2 4 |
| Inverse | -2 1 / 1.5 -0.5 |
| Extract diagonal | 1 / 4 |
| Scale (k=2) | 2 4 / 6 8 |
| Power (p=2) | 7 10 / 15 22 |

**Multiply / Multiplication-Hadamard / Multiplication-Kronecker / Add** (and other two-matrix ops): first InputBox is matrix A (defaults to the selection), second is matrix B. Multiply A1:B2 by itself → 7 10 / 15 22. **Multiplication-Hadamard** is the element-wise product; if A and B differ in shape, an OK message box says they must be the same shape and nothing is written. If they match, a third InputBox asks for the output cell (default one column to the right of B). **Multiplication-Kronecker** is the Kronecker product (A ⊗ B).

**Dot product** of `{1;2}` and `{3;4}` → 11. **Outer product** of the same → 3 4 / 6 8.

Cancel on either InputBox must not write.

## Properties

Same A1:B2. **Matrices → Properties**. Each item prompts for the matrix and an output cell (default two rows below, same as Size) and writes the property name with the value in the cell to its right.

| Menu | Label | Value |
|------|-------|-------|
| All properties | eight scalar rows, then Eigenvalues and Eigenvectors | includes Condition number and Spectral radius |
| Determinant | Determinant | -2 |
| Trace | Trace | 5 |
| Rank | Rank | 2 |
| Frobenius norm | Frobenius norm | sqrt(30) |
| 1-norm | 1-norm | 6 |
| Infinity-norm | Infinity-norm | 7 |
| Condition number | Condition number | 21 (∞-norm, ||A||_∞ ||A⁻¹||_∞) |
| Spectral radius | Spectral radius | 5.372 (max \|λ\|; exact when symmetric) |
| Eigenvalues | Eigenvalues | column of λ to the right (symmetric matrices; else "Needs a symmetric matrix") |
| Eigenvectors | Eigenvectors | n × n matrix to the right of the label (columns are eigenvectors) |

## Validation

| Menu | Expected |
|------|----------|
| Is symmetric | `Symmetric` and TRUE/FALSE in the cell to its right (prompts for the matrix and an output cell; default is two rows below, same as Size) |
| Hadamard Proof | Prompts for the matrix. If any entry is not 1 or -1, a message box says the input is not Hadamard and no sheet is written. Otherwise writes **Hadamard Proof**: H, Hᵀ, H.HT. Order-2 Hadamard `1 1 / 1 -1` → n = 2 and I₂. A ±1 matrix that fails H Hᵀ = n I → `Not Hadamard: H.HT not equal to nI` |

## Decompositions

- **Cholesky** on a SPD matrix such as 4 2 / 2 3. Lower L to the right; `L * L^T` recovers A.
- **Eigen (symmetric)** on 2 1 / 1 2. Last column eigenvalues near 3 and 1; A v ≈ λ v.
- **QR** on a tall or square numeric range. Q on top, R immediately below; Q^T Q ≈ I and Q R ≈ A.
- **LU** on 4 2 / 2 3 (no row swap). L on top, U below; L U ≈ A. On a matrix that needs pivoting, L U equals P A rather than A.

## Worksheet UDFs

After restart, Insert Function should list a category **Excel VBA Lib** with `Mat*` names. Array-enter or use dynamic arrays:

```
=MatMult(A1:B2,A1:B2)
=MatInv(A1:B2)
=MatDet(A1:B2)
=MatChol(A1:B2)
=MatLU(A1:B2)
=MatRank(A1:B2)
=MatScale(A1:B2,2)
=MatrixMultDefined(A1:B2,A1:B2)
```

Empty or text cells → `#VALUE!`. Singular inverse / non-SPD Cholesky / non-symmetric Eigen → `#NUM!` or `#VALUE!` (UDF) or an error message (menu).

**Cofactor** on 1 2 / 3 4 → 4 -3 / -2 1. **Minor** deleting row 1 column 1 → 4.

In the add-in VBE (`Alt+F11` on `ExcelVbaLib.xlam`) you should see `modApiMatrices1` and `modApiMatrices2`. If those modules are still the short git wrappers, run `Import-Menu13FromPersonal.ps1` to overlay Personal (names are rewritten to `modApi*`).

## Limits

Side length above 250 is rejected. Results write one column to the right of the selection (create writes at the active cell). Complex/unitary helpers from Personal `custom_Menu13_Unitary` are not in this pack.
