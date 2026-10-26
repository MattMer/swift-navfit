param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$FixturePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $FixturePath){. (Join-Path $PSScriptRoot 'test-pdf-fixtures.ps1');$FixturePath=NewSwiftPdfFixtures 'expanded'}
$FixturePath=[IO.Path]::GetFullPath($FixturePath)
$scratch=Join-Path $root ('build/tests/pdf-expanded-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$hashes=@{};Get-ChildItem $FixturePath -Filter '*.pdf' | ForEach-Object {$hashes[$_.FullName]=(Get-FileHash -LiteralPath $_.FullName).Hash}
$app=$null
try{
    $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'PDFExpandedHarness',(Join-Path $root 'tests/PDFExpandedHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestExpandedPDF("'+$FixturePath+'")')
    if(-not $result.StartsWith('PASS:')){throw $result}
    foreach($path in $hashes.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $hashes[$path]){throw 'Test changed fixture'}}
    Write-Output $result
}finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
