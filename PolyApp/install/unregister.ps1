<#
.SYNOPSIS
  Remove the PolyApp add-in registration (per-user keys; -Machine for HKLM via elevated RegAsm /u and the HKLM add-in key).
#>
param([switch]$Machine, [string]$Configuration = "Release")

$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
$dll  = Join-Path $env:LOCALAPPDATA "PolyApp\AddIn\PolyApp.AddIn.dll"
if (-not (Test-Path $dll)) { $dll = Join-Path $root "AddIn\bin\$Configuration\net48\PolyApp.AddIn.dll" }
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$guid = "{5C1D7A2E-9B64-4F3A-8E21-7D0A3B9C4E55}"

if ($Machine) {
    Start-Process -FilePath $regasm -ArgumentList "`"$dll`" /unregister" -Verb RunAs -Wait | Out-Null
    $keyScript = Join-Path $PSScriptRoot "hklm-addin-key.ps1"
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$keyScript`" -Remove" -Verb RunAs -Wait | Out-Null
}
foreach ($k in @("HKCU:\SOFTWARE\SolidWorks\AddIns\$guid", "HKCU:\SOFTWARE\SolidWorks\AddInsStartup\$guid",
                 "HKCU:\Software\Classes\CLSID\$guid", "HKCU:\Software\Classes\PolyApp.AddIn")) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Host "removed $k" }
}
Write-Host "done (the HKLM add-in key needs -Machine)"
