param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/major-ui-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    foreach($module in @('MajorHarness','MajorUIHarness')){$app.LoadFromText(5,$module,(Join-Path $root ('tests/'+$module+'.bas')))}
    $result=$app.Eval('TestMajor()');if(-not $result.StartsWith('PASS:')){throw $result}
    $result=$app.Eval('TestMajorUI()');if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
    $result=$app.Eval('TestMajorViewBounds()');Write-Output $result
    if(-not $result.StartsWith('PASS:')){throw $result}
}finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
