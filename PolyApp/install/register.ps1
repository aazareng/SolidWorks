<#
.SYNOPSIS
  Build and register the PolyApp add-in with SOLIDWORKS.

.DESCRIPTION
  Per-user by default (no admin): RegAsm writes a .reg file, which is rewritten from
  HKEY_CLASSES_ROOT to HKCU\Software\Classes and extended with the SolidWorks AddIns keys
  under HKCU, then imported. SOLIDWORKS 2026 lists add-ins from HKLM only, so that one key
  is written through a UAC prompt (install\hklm-addin-key.ps1).

  -Machine registers for all users (RegAsm /codebase, elevated prompt).
  -NoBuild skips `dotnet build`.
  -Load asks the running SOLIDWORKS to load the add-in right away (no restart) via the polyapp CLI.

.EXAMPLE
  .\install\register.ps1
  .\install\register.ps1 -Load
  .\install\register.ps1 -Machine
#>
param([switch]$Machine, [switch]$NoBuild, [switch]$Load, [switch]$NoHklm, [string]$Configuration = "Release")

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$proj = Join-Path $root "AddIn\PolyApp.AddIn.csproj"
$dll  = Join-Path $root "AddIn\bin\$Configuration\net48\PolyApp.AddIn.dll"
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$guid = "{5C1D7A2E-9B64-4F3A-8E21-7D0A3B9C4E55}"

if (-not $NoBuild) {
    Write-Host "building $proj"
    & dotnet build $proj -c $Configuration -nologo -v q
    if ($LASTEXITCODE -ne 0) { throw "build failed" }
}
if (-not (Test-Path $dll)) { throw "add-in DLL not found: $dll" }
if (-not (Test-Path $regasm)) { throw "RegAsm not found: $regasm" }

# Deploy a copy: SOLIDWORKS keeps a loaded add-in DLL locked, so the registered copy lives apart
# from the build output and `dotnet build` keeps working while SOLIDWORKS is open.
# The deployed PolyApp.ini is the SHIPPED one; per-user edits belong in %APPDATA%\PolyApp\PolyApp.ini.
$deployRoot = Join-Path $env:LOCALAPPDATA "PolyApp"
$deploy = Join-Path $deployRoot "AddIn"
$src = Split-Path -Parent $dll
$leaf = Split-Path -Leaf $dll
$locked = $false
if (Test-Path (Join-Path $deploy $leaf)) {
    try { $fs = [IO.File]::Open((Join-Path $deploy $leaf), 'Open', 'ReadWrite', 'None'); $fs.Close() } catch { $locked = $true }
}
if ($locked) {
    # SOLIDWORKS is holding the current copy: deploy beside it and point the registration at the new folder.
    $deploy = Join-Path $deployRoot ("AddIn-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
    Write-Warning "the registered copy is in use by SOLIDWORKS; deploying to $deploy instead. Restart SOLIDWORKS to pick it up."
}
New-Item -ItemType Directory -Force $deploy | Out-Null
& robocopy.exe $src $deploy /MIR /R:2 /W:1 /NJH /NJS /NFL /NDL /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw "could not copy the add-in to $deploy" }
# old versioned copies that are no longer locked
Get-ChildItem $deployRoot -Directory -Filter "AddIn-*" | Where-Object { $_.FullName -ne $deploy } | ForEach-Object {
    try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
}
$dll = Join-Path $deploy $leaf

if ($Machine) {
    Write-Host "registering for all users (elevated RegAsm /codebase)"
    $p = Start-Process -FilePath $regasm -ArgumentList "`"$dll`" /codebase" -Verb RunAs -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw "RegAsm exit code $($p.ExitCode)" }
} else {
    $regfile = Join-Path $env:TEMP "PolyApp.AddIn.reg"
    & $regasm "$dll" /codebase /regfile:"$regfile" | Out-Null
    if (-not (Test-Path $regfile)) { throw "RegAsm did not produce $regfile" }
    $txt = Get-Content $regfile -Raw
    $txt = $txt -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes"
    $txt += @"

[HKEY_CURRENT_USER\SOFTWARE\SolidWorks\AddIns\$guid]
@=dword:00000000
"Title"="PolyApp"
"Description"="PolyApp engineering commands: PDM state, Branch, Open, Make Drawing, Hotsheet, Shipping Order, Cutsheet, Power Select, exports."

[HKEY_CURRENT_USER\SOFTWARE\SolidWorks\AddInsStartup\$guid]
@=dword:00000001

"@
    Set-Content -Path $regfile -Value $txt -Encoding Unicode
    Write-Host "importing $regfile (per-user)"
    # reg.exe reports success on stderr, which PowerShell 5.1 turns into an error under -ErrorActionPreference Stop: run it detached.
    $p = Start-Process -FilePath reg.exe -ArgumentList "import `"$regfile`"" -Wait -PassThru -WindowStyle Hidden
    if ($p.ExitCode -ne 0) { throw "reg import failed ($($p.ExitCode))" }

    # SOLIDWORKS 2026 lists add-ins from HKLM\SOFTWARE\SolidWorks\AddIns only (the HKCU copy is ignored
    # by the Add-Ins dialog), so that one key is written elevated. The COM class stays per-user.
    # -NoHklm skips the prompt (the add-in still loads through `polyapp load-addin` / -Load; it is just not listed in the dialog).
    $hklm = "HKLM:\SOFTWARE\SolidWorks\AddIns\$guid"
    if ($NoHklm) { Write-Host "skipping $hklm (-NoHklm)" }
    elseif (-not (Test-Path $hklm)) {
        Write-Host "writing $hklm (UAC prompt)"
        $keyScript = Join-Path $PSScriptRoot "hklm-addin-key.ps1"
        $p = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$keyScript`"" -Verb RunAs -Wait -PassThru
        if ($p.ExitCode -ne 0) { Write-Warning "could not write $hklm (exit $($p.ExitCode)); SOLIDWORKS will not list the add-in" }
    }
}

Write-Host "registered $dll"
Write-Host "In SOLIDWORKS: Tools > Add-Ins > 'PolyApp' (a restart picks it up automatically). Log: $env:LOCALAPPDATA\PolyApp\PolyApp.log"

if ($Load) {
    $cli = Join-Path $root "Cli\bin\$Configuration\net10.0-windows\polyapp.exe"
    if (-not (Test-Path $cli)) { & dotnet build (Join-Path $root "Cli\polyapp.csproj") -c $Configuration -nologo -v q }
    & $cli load-addin "$dll"
}
