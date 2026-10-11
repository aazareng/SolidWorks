# Writes the machine-wide SolidWorks add-in key (needs elevation). Called by register.ps1 via UAC.
param([switch]$Remove)
$k = "HKLM:\SOFTWARE\SolidWorks\AddIns\{5C1D7A2E-9B64-4F3A-8E21-7D0A3B9C4E55}"
if ($Remove) { if (Test-Path $k) { Remove-Item $k -Recurse -Force }; exit 0 }
New-Item -Path $k -Force | Out-Null
Set-ItemProperty -Path $k -Name "(default)" -Value 0 -Type DWord
Set-ItemProperty -Path $k -Name "Title" -Value "PolyApp"
Set-ItemProperty -Path $k -Name "Description" -Value "PolyApp engineering commands: PDM state, Branch, Open, Make Drawing, Hotsheet, Shipping Order, Cutsheet, Power Select, exports."
exit 0
