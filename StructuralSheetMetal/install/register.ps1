<#
.SYNOPSIS
  Build and register the Structural Sheet Metal add-in with SOLIDWORKS.

.DESCRIPTION
  Per-user by default (no admin): RegAsm writes a .reg file, which is rewritten from
  HKEY_CLASSES_ROOT to HKCU\Software\Classes and extended with the SolidWorks AddIns keys
  under HKCU, then imported. SOLIDWORKS reads add-ins from both HKLM and HKCU.

  -Machine registers for all users (RegAsm /codebase, elevated prompt).
  -NoBuild skips `dotnet build`.
  -Load asks the running SOLIDWORKS to load the add-in right away (no restart) via the ssm CLI.

.EXAMPLE
  .\install\register.ps1
  .\install\register.ps1 -Load
  .\install\register.ps1 -Machine
#>
param([switch]$Machine, [switch]$NoBuild, [switch]$Load, [string]$Configuration = "Release")

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$proj = Join-Path $root "AddIn\StructuralSheetMetal.AddIn.csproj"
$dll  = Join-Path $root "AddIn\bin\$Configuration\net48\StructuralSheetMetal.AddIn.dll"
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$guid = "{8F2B6C3A-5D41-4E7B-9A10-2C6F4B1E7D30}"

if (-not $NoBuild) {
    Write-Host "building $proj"
    & dotnet build $proj -c $Configuration -nologo -v q
    if ($LASTEXITCODE -ne 0) { throw "build failed" }
}
if (-not (Test-Path $dll)) { throw "add-in DLL not found: $dll" }
if (-not (Test-Path $regasm)) { throw "RegAsm not found: $regasm" }

# Deploy a copy: SOLIDWORKS keeps a loaded add-in DLL locked, so the registered copy lives apart
# from the build output and `dotnet build` keeps working while SOLIDWORKS is open.
$deployRoot = Join-Path $env:LOCALAPPDATA "StructuralSheetMetal"
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
    $regfile = Join-Path $env:TEMP "StructuralSheetMetal.AddIn.reg"
    & $regasm "$dll" /codebase /regfile:"$regfile" | Out-Null
    if (-not (Test-Path $regfile)) { throw "RegAsm did not produce $regfile" }
    $txt = Get-Content $regfile -Raw
    $txt = $txt -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes"
    $txt += @"

[HKEY_CURRENT_USER\SOFTWARE\SolidWorks\AddIns\$guid]
@=dword:00000000
"Title"="Structural Sheet Metal"
"Description"="Sheet-metal members from open profiles along a sketch path, with corner treatments."

[HKEY_CURRENT_USER\SOFTWARE\SolidWorks\AddInsStartup\$guid]
@=dword:00000001

"@
    Set-Content -Path $regfile -Value $txt -Encoding Unicode
    Write-Host "importing $regfile (per-user)"
    & reg.exe import "$regfile" 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "reg import failed ($LASTEXITCODE)" }
}

Write-Host "registered $dll"
Write-Host "In SOLIDWORKS: Tools > Add-Ins > 'Structural Sheet Metal' (a restart picks it up automatically)."

if ($Load) {
    $ssm = Join-Path $root "Cli\bin\$Configuration\net10.0-windows\ssm.exe"
    if (-not (Test-Path $ssm)) { & dotnet build (Join-Path $root "Cli\ssm.csproj") -c $Configuration -nologo -v q }
    & $ssm load-addin "$dll"
}
