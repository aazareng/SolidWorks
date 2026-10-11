<#
.SYNOPSIS
  Remove the Structural Sheet Metal add-in registration (per-user keys; -Machine for HKLM via elevated RegAsm /u).
#>
param([switch]$Machine, [string]$Configuration = "Release")

$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
$dll  = Join-Path $root "AddIn\bin\$Configuration\net48\StructuralSheetMetal.AddIn.dll"
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$guid = "{8F2B6C3A-5D41-4E7B-9A10-2C6F4B1E7D30}"

if ($Machine) {
    Start-Process -FilePath $regasm -ArgumentList "`"$dll`" /unregister" -Verb RunAs -Wait | Out-Null
}
foreach ($k in @("HKCU:\SOFTWARE\SolidWorks\AddIns\$guid", "HKCU:\SOFTWARE\SolidWorks\AddInsStartup\$guid",
                 "HKCU:\Software\Classes\CLSID\$guid", "HKCU:\Software\Classes\StructuralSheetMetal.AddIn")) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Host "removed $k" }
}
Write-Host "done"
