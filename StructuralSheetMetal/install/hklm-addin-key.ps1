# Writes the machine-wide SolidWorks add-in key (needs elevation). Called by register.ps1 via UAC.
param([switch]$Remove)
$k = "HKLM:\SOFTWARE\SolidWorks\AddIns\{8F2B6C3A-5D41-4E7B-9A10-2C6F4B1E7D30}"
if ($Remove) { if (Test-Path $k) { Remove-Item $k -Recurse -Force }; exit 0 }
New-Item -Path $k -Force | Out-Null
Set-ItemProperty -Path $k -Name "(default)" -Value 0 -Type DWord
Set-ItemProperty -Path $k -Name "Title" -Value "Structural Sheet Metal"
Set-ItemProperty -Path $k -Name "Description" -Value "Sheet-metal members from open profiles along a sketch path, with corner treatments."
exit 0
