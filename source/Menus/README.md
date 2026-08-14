# Menus

Ribbon / command-bar construction and add-in workbook events.

| Personal | Role |
|----------|------|
| `modAddinMenu` | Add-in `Auto_Open` menu (**Excel VBA Lib**) so callers need not import modules |
| `Custom_Menu_Menus` | Personal ribbon (OnAction names must match Api/Feature `Public Sub`s) |
| `Custom_Menu12_Commandbar` | command-bar helpers |
| `ThisWorkbook` | `Workbook_Open` — stays in the `.xlam` document module; keep a `.bas` copy here for git only |

Do not treat menu captions as the API. The API is the `Public Sub` each `.OnAction` points at.
