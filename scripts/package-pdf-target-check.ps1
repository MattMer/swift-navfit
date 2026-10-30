param([string]$OutputPath=(Join-Path $PSScriptRoot '..\build\pdf-target-check\Swift_NAVFIT_Local_PDF_Check.zip'),
 [string]$FixtureDirectory)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $FixtureDirectory){. (Join-Path $PSScriptRoot 'test-pdf-fixtures.ps1');$FixtureDirectory=NewSwiftPdfFixtures 'expanded'}
$OutputPath=[IO.Path]::GetFullPath($OutputPath);[void](New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath))
if(Test-Path -LiteralPath $OutputPath){throw "Refusing to overwrite $OutputPath"}
$stage=Join-Path $root ('build/tests/pdf-target-kit-'+[guid]::NewGuid())
foreach($directory in @('scripts','scripts/cases','tests','output/pdf/r126')){[void](New-Item -ItemType Directory -Force -Path (Join-Path $stage $directory))}
foreach($script in @('test-pdf-output.ps1','suite-common.ps1')){Copy-Item -LiteralPath (Join-Path $PSScriptRoot $script) -Destination (Join-Path $stage 'scripts')}
Copy-Item -LiteralPath (Join-Path $root 'scripts/cases/test-pdf-target-computer.ps1') -Destination (Join-Path $stage 'scripts/cases')
Copy-Item -LiteralPath (Join-Path $root 'tests/PDFDependencyHarness.bas') -Destination (Join-Path $stage 'tests')
foreach($kind in @('eval','chiefeval','fitrep')){Copy-Item -LiteralPath (Join-Path $FixtureDirectory ($kind+'-expanded.pdf')) -Destination (Join-Path $stage 'output/pdf/r126')}
@'
Swift NAVFIT r126 local PDF extraction acceptance check

Unzip this kit, open PowerShell in this folder, then run:

powershell.exe -NoProfile -File .\scripts\test-pdf-output.ps1 -TargetComputerOnly -BuildPath 'C:\YourFolder\Swift_NAVFIT_Blank_2026-10-20-r126.accdb'

Use r126 or newer and desktop Microsoft Access. Extraction is embedded in the
Access application. Acrobat, Python, external PDF utilities and an internet
connection are not needed. No installation or security setting is changed.

The three PDFs contain synthetic information only. The script checks their
mapped fields using a temporary copy of the app; it does not import reports into
the supplied app. A JSON report under output\pdf records Access build/bitness,
input hashes and per-form results. Original app/PDF hashes must be unchanged.
Use -ReportPath for another writable report location. Failure exits with code 1.

Run under the intended user's normal approved configuration. PASS verifies this
extraction path on that computer. Physical display, accessibility and printing
acceptance remain separate. The in-app Check PDF extraction on this computer
button can also check a selected supported May 2025 fillable NAVPERS form.
'@ | Set-Content -LiteralPath (Join-Path $stage 'README.txt') -Encoding UTF8
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $OutputPath
Write-Output "Created: $OutputPath"
