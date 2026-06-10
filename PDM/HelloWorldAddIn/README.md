# HelloWorldAddIn

A minimal SOLIDWORKS PDM Professional add-in (.NET, C#). Right-click on a
file in a vault view to get two new commands:

- **Hello World** - shows a message box (sanity check that the add-in loads).
- **Set to Issued** - calls `IEdmFile5.ChangeState(...)` on the selected
  file(s), the same way `PDMtoState` / `ChangeState` does in the VBA macro.

## Build

1. Open this folder in Visual Studio (or `dotnet build` with the .NET 4.8
   targeting pack installed).
2. Fix the `HintPath` in `HelloWorldAddIn.csproj` to point at your local
   `EPDM.Interop.epdm.dll`, normally:
   `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS PDM\EPDM.Interop.epdm.dll`
3. Build in **x64**, Release configuration. Because
   `RegisterForComInterop` is set, Visual Studio (run as Administrator)
   will register the COM class automatically on build.

If you build via `dotnet build` instead of full Visual Studio, register the
DLL manually (run as Administrator):

```
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" HelloWorldAddIn.dll /codebase
```

## Deploy

1. Open the **PDM Administration** tool, connect to your vault.
2. Right-click **Add-ins** -> **New Add-in...**.
3. Browse to `HelloWorldAddIn.dll` (built output). PDM reads the GUID from
   the registry/manifest and stores the add-in in the vault database, then
   pushes it to clients on next sync.
4. In the **Add-in Settings**, make sure the add-in is enabled for the
   users/groups you want to test with.
5. On a client, right-click a file in the vault view - you should see
   "Hello World" and "Set to Issued" in the context menu.

## Customizing the state change

`TargetStateName` in `AddIn.cs` is hardcoded to `"Issued"`. Change it (or
make it driven by the current state, like the VBA `PDMtoState`/`SetppPDM`
logic) to match your workflow's transition names. `ChangeState` requires the
transition from the file's *current* state to `TargetStateName` to exist in
the workflow, otherwise it will throw - same as the VBA version.
