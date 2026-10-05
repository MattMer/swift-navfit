param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build\tests\major-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1
    $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'MajorHarness',(Join-Path $root 'tests\MajorHarness.bas'))
    $result=$app.Eval('TestMajor()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
} finally {
    if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
