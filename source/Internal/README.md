# Internal

Shared helpers used by Api and Features in the **same** `.xlam`. Do not document these as the external API. **Internal must not call Api** (avoids cycles).

| File | Role |
|------|------|
| `modInternalExcelApp.bas` | `PushAppState` / `PopAppState` (screen, calc, events) |
| `modInternalSheetIO.bas` | array ↔ range |
| `modInternalError.bas` | `RaiseCurrent` |
| `modInternalText.bas` | string / blank helpers |
| `modInternalNamedRanges.bas` | create / check names |
| `modInternalDateTable.bas` | calendar dimension array |
| `modInternalBenford.bas` | Benford table + charts |

Candidates still in Personal `aPublic*` (see [docs/InfrastructureCatalog.md](../../docs/InfrastructureCatalog.md)):

- `modInternalRangeValid` ← `aPublic_RangeValid`
- `modInternalRangeFn` ← `aPublic_RangeFn`
- `modInternalArrayFn` / `modInternalArrayValid`
- `modInternalFormat` ← `aPublic_Format`
