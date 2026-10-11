# PolyApp — a SOLIDWORKS add-in for the PolyAppPro commands

The PolyAppPro macro (`PolyAppPro.swp`, 56 modules, ~640 procedures, 38 actions on one form) as a real
add-in: a **PolyApp** tab in parts, assemblies and drawings, flyouts for the command families, entries on
the right-click menu of a component, one log, one ini. Nothing is rewritten up front: every button runs
the existing macro through a named entry point, and a command moves to C# when it is next touched.

Status, 2026-10-11, written on the home PC (SOLIDWORKS 2026 SP2.1, no vault, no `.swp`):

| Piece | State |
|---|---|
| `Core\` settings, catalogue, dispatcher, macro bridge | Builds; `polyapp ini` / `polyapp commands` verified from the terminal |
| `AddIn\` tab, flyouts, context menu, icons, events | Loaded live into SOLIDWORKS 2026 SP2.1 via `register.ps1 -NoHklm -Load`: ConnectToSW ran, 18 buttons + 6 flyouts created, no errors. The tab itself and the buttons are **not yet eyeballed or clicked** |
| `macro\mod_PolyAppEntry.bas` | **UNTESTED**: needs the work `.swp` and three small edits (described at the top of the file) |
| Native commands | `OpenDrawing` (drawing beside the model, else the macro), `SelectTree`, `Settings`, `ShowLog` |

## How an add-in works, in one paragraph

A macro is a script you run. An add-in is a DLL that SOLIDWORKS loads at startup: a registry key under
`SOFTWARE\SolidWorks\AddIns` names its COM class, SOLIDWORKS creates it and calls `ConnectToSW`, and from
then on the code lives inside the SOLIDWORKS process like Extrude does. That lets it own a CommandManager
tab, flyout buttons, PropertyManager pages, right-click menu entries and event hooks, with buttons that
enable and disable by document type. The DLL must target .NET Framework 4.8 (SOLIDWORKS loads add-ins
through mscoree). The engine lives in `Core\` and is compiled into both the add-in and `polyapp.exe`, so
everything is testable from a terminal without registering anything.

## Layout

```
PolyApp\
  PolyApp.ini            every location and option the macro had as a Global Const (vault, server, templates, PDM, e-mail, API)
  Core\                  shared, compiled into BOTH projects
    Settings.cs            layered ini: defaults -> PolyApp.ini beside the DLL -> %APPDATA%\PolyApp\PolyApp.ini -> POLYAPP_INI
    Commands.cs            THE CATALOGUE: id, title, box, flyout, document types, frm_Main control(s), context-menu flag
    Dispatcher.cs          runs a command: native handler first, the macro otherwise (a handler may return null = "not my case")
    MacroRunner.cs         writes %TEMP%\PolyApp.request.ini, calls RunMacro2 on PA_Run
    DocContext.cs          active document kind, path, selection
    PdmSession.cs          one vault login at startup (compiled with EPDM.Interop.epdm when the PDM client is installed)
    Native\OpenDrawing.cs  the first native command and the pattern
    Native\Setup.cs        Settings (opens your ini) and Log buttons
  AddIn\                 net48 COM add-in: SwAddin.cs (ISwAddin, tab, flyouts, popup menu, events, Cmd_*/Enable_* callbacks, registration)
                         IconFactory.cs (icon strips generated from the catalogue; drop <Id>.png in AddIn\Icons to override)
  Cli\                   net10 console: polyapp.exe (attach | ini | commands | run <Id> | load-addin | unload-addin)
  macro\mod_PolyAppEntry.bas   import into PolyAppPro.swp: PA_Run reads the request and presses Run on frm_Main
  install\register.ps1   build + deploy to %LOCALAPPDATA%\PolyApp\AddIn + register (per-user; -Machine for HKLM; -Load = no restart)
  install\unregister.ps1
```

## Build, register, test

```powershell
dotnet build Cli\polyapp.csproj -c Release
$pa = "Cli\bin\Release\net10.0-windows\polyapp.exe"
& $pa ini                      # resolved settings and which file each came from
& $pa commands                 # the catalogue with its native/macro mapping
& $pa attach                   # SOLIDWORKS, active document, PDM status, macro found or not
& $pa run OpenDrawing          # one command, exactly as the button would run it

.\install\register.ps1         # builds, deploys, registers per-user (one UAC prompt for the HKLM listing key)
# restart SOLIDWORKS -> Tools > Add-Ins > PolyApp -> the PolyApp tab
```

Log: `%LOCALAPPDATA%\PolyApp\PolyApp.log` (the Log button on the tab opens it).

## The tab

| Box | Buttons | Flyouts |
|---|---|---|
| PDM | Branch, Open, Open Drawing | **PDM State**: Check Out, Check In, Issue, Issue Next Rev, Banish, Obsolete |
| Make | Make Drawing, Cutsheet, Mate Origins, Mirrorize | **Make Assembly**: Hotsheet, Shipping Order, Shipping (SH), POR, Manual, Spec, Cutsheet Assembly, Modified (-M) |
| Select | Tree (toggle, shows pressed) | **By Type**, **By Department**, **By State**: the 27 Power Select options |
| Export | DXF, PDF, eDrawings, STEP, Export BOM, Export HTML | |
| Drawing | Reload Format, Check Balloons | |
| More | | **More**: Add/Sync Config, Configger, Colorize, Pinkify, Purge Configs, Trim Properties/Configs, Cut Size, Set Material, CSV Assembly, Error Check, Replace Ref, Dissolve, Exclude from BOM, Add Spec, Sync Card, Delete Document |
| Setup | Settings, Log | |

Each tab only shows the buttons that apply to that document type (the `_MDL` / `_ASM` / `_GEN` / `_DWG`
suffixes of the form). Right-click on a component: Check Out, Check In, Branch, Open, Open Drawing, Hotsheet.

## How a button runs

1. `Cmd_<Id>` in `SwAddin.cs` (SOLIDWORKS calls it by name) → `Dispatcher.Run(id)`.
2. A native handler registered for the id runs first. It returns a result, or `null` to hand over to the macro.
3. Otherwise `MacroRunner` writes `%TEMP%\PolyApp.request.ini` with the frm_Main control(s) from the catalogue
   (overridable per command in `[MacroOptions]`) and calls `RunMacro2` on `mod_PolyAppEntry.PA_Run`.
4. `PA_Run` loads frm_Main (its Initialize fills the serial number from the active document), sets the
   control(s) and the Tree box, and presses Run.

So the work `.swp` stays the source of truth. Porting a command = one native handler in `Core\Native\`
plus one line in `Dispatcher.BuiltIn`; the macro path stays as the fallback until the port covers every case.

## Make Assembly

The own commands sit beside it (the Make box). To put the PolyApp menu *in place of* SOLIDWORKS' Make
Assembly: set `[Debug] LogCommands=yes`, restart, click Make Assembly once, read its id from the log, put it
in `[Intercept] MakeAssembly=`. The add-in then cancels that command in parts and assemblies and pops up
Hotsheet / Shipping Order / ... / Make Assembly (SOLIDWORKS) at the cursor.

## At work: first-day checklist

1. Copy the repo; `dotnet build` both projects (needs the .NET 10 SDK; the PDM interop is picked up from
   the PDM client install automatically).
2. In the vault's `PolyAppPro.swp`: the three edits at the top of `macro\mod_PolyAppEntry.bas`, import the module, save.
3. `polyapp commands` → compare the control names against the work form; fix any that differ in
   `%APPDATA%\PolyApp\PolyApp.ini` under `[MacroOptions]` (the `Modified` command has no mapping yet).
4. `polyapp run Open` from a terminal with a part open: the first end-to-end macro call.
5. `install\register.ps1 -Load`, then the tab.
