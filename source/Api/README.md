# Api

Public entry points for the add-in. Other workbooks and the ribbon should call **only** these `Public` names.

| File | Personal origin / role |
|------|------------------------|
| `modApiArrays.bas` | `aPublicProcedures` array dump/load |
| `modApiSheets.bas` | output / ensure / clear sheets |
| `modApiFiles.bas` | folders, file lists |
| `modApiTables.bas` | ListObjects |
| `modApiDates.bas` | `CreateDateTable` (`Custom_Menu23_DateTable`) |
| `modApiBenford.bas` | Benford analyses (`Custom_Menu5_Benford*`) |
| `modApiUi.bas` | status / message wrappers |

Add a new `modApi{Area}.bas` when a Personal family has a stable public surface. Heavy work belongs in `../Internal/`.
