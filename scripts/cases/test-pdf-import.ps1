param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'test-pdf-fixtures.ps1');$pdfs=NewSwiftPdfFixtures 'synthetic'
$scratch=Join-Path $root ('build\tests\pdf-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'PDFHarness',(Join-Path $root 'tests\PDFHarness.bas'))
    $result=$app.Eval('TestPDFImport("'+$pdfs+'")')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
} finally {
    if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
