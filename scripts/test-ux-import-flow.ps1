param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$SkipPDF)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-import-flow-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$navfit=Join-Path $root ('build/tests/ux-import-flow-source-'+[guid]::NewGuid()+'.accdb'); & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $navfit
$sourceHash=(Get-FileHash -LiteralPath $navfit).Hash
$pdfHash=$null;if(-not $SkipPDF){$pdfHash=(Get-FileHash -LiteralPath (Join-Path $root 'output/pdf/synthetic-eval.pdf')).Hash}
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXImportFlowHarness',(Join-Path $root 'tests/UXImportFlowHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestUXImportFlow("'+$navfit+'","'+(Join-Path $root 'output/pdf/synthetic-eval.pdf')+'",'+$(if($SkipPDF){'False'}else{'True'})+')')
    if(-not $result.StartsWith('PASS:')){throw $result}
    if((Get-FileHash -LiteralPath $navfit).Hash -ne $sourceHash -or (-not $SkipPDF -and (Get-FileHash -LiteralPath (Join-Path $root 'output/pdf/synthetic-eval.pdf')).Hash -ne $pdfHash)){throw 'Import flow checks changed a source file'}
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
