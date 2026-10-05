param([string]$OutputPath='D:\Development\Swift NAVFIT\output\pdf\Swift_NAVFIT_PDF_Dependency_Check_r26.zip')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw "Refusing to overwrite $OutputPath"}
$stage=Join-Path $root ('build/tests/pdf-target-kit-'+[guid]::NewGuid())
foreach($directory in @('scripts','tests','output/pdf/r26')){[void](New-Item -ItemType Directory -Force -Path (Join-Path $stage $directory))}
Copy-Item -LiteralPath (Join-Path $root 'scripts/test-pdf-target-computer.ps1') -Destination (Join-Path $stage 'scripts')
Copy-Item -LiteralPath (Join-Path $root 'tests/PDFDependencyHarness.bas') -Destination (Join-Path $stage 'tests')
foreach($kind in @('eval','chiefeval','fitrep')){Copy-Item -LiteralPath (Join-Path $root ('output/pdf/r26/'+$kind+'-expanded.pdf')) -Destination (Join-Path $stage 'output/pdf/r26')}
@'
Swift NAVFIT r26 PDF extraction acceptance check

Unzip this kit, open PowerShell in this folder, then run:

powershell.exe -NoProfile -File .\scripts\test-pdf-target-computer.ps1 -BuildPath 'C:\YourFolder\Swift_NAVFIT_UI_Demo_2026-10-04-r26.accdb'

Use an r26 app (blank or demo), desktop Microsoft Access and desktop Adobe Acrobat.
Reader/browser PDF viewers do not supply the required automation. No software
installation, licensing, trust or security configuration is changed by this kit.

The three PDFs contain synthetic information only. This reads their structured
data, verifies the expected values and uses a temporary copy of the app. It never
imports reports into the supplied app. A JSON report appears under output\pdf
with Access build/bitness, Acrobat version, input hashes and per-form results.
The original app/PDF hashes must remain unchanged. You can set -ReportPath to a
writable location. A failing check exits with code 1; keep its report for diagnosis.

Run as the intended user under the computer's normal approved configuration.
PASS covers this extraction path on this computer, not complete workstation,
Access Runtime, display/accessibility or printing acceptance. Runtime users can
use the app's native Check PDF extraction on this computer button with a chosen
supported PDF; this script additionally requires Access module-import automation.
'@ | Set-Content -LiteralPath (Join-Path $stage 'README.txt') -Encoding UTF8
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $OutputPath
Write-Output "Created: $OutputPath"
