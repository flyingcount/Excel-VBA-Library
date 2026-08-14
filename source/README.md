# source/

These `.bas` files are a snapshot of what was imported into **`build/ExcelVbaLib.xlam`**.

**The add-in is the library.** Add new modules in the VBE of `ExcelVbaLib.xlam` (`Alt+F11` → Insert → Module). Do not create new git folders here for Personal123 families.

`_export_raw/` is a local dump from `Data/Personal123.xlsb` (gitignored). It is not imported into the add-in.

To rebuild the current add-in from this snapshot:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Build-ExcelVbaLib.ps1
```
