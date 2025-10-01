# ABMEP – Manifest

**Revit:** 2024  
**.NET:** .NET Framework 4.8  

## Ribbon
- Tab: `ABMEP`
- Loader: `ABMEP/App.cs` (CSV-driven)
- CSV: `ABMEP.tools.csv` supports 4/6/8 cols  
  - 8-col adds `Size,StackKey` (use `small,SyncA` to stack)
- Icons: 16/32 PNGs next to the DLL (or `icons/`)

## Solutions / Projects
- `ABMEP.Tools` – Hotload launchers (derive `HotloadCommandBase`)
- `ABMEP.Work` – Commands (`IExternalCommand`)

## Conventions
- Exports go to `C:\Temp`
- Prefer complete replacements over patches

## Parameters (current tools)
- Read: `Product Entry`, `eM_Width`, `eM_Length A`, `eM_Length B`
- Write: `Hanger Size`, `Strut Length`, `Rod Length`

## Tools (examples)
- Reports → Hanger Inserts (CSV)
- Parameter Sync → Hanger Size from Product Entry
- Parameter Sync → Strut Length from eM_Width
- Parameter Sync → Rod Length from eM_Length A+B
- Trimble → Hangers/Sleeves
