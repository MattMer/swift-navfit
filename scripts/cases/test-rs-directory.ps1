param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/rs-directory-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'RSDirectoryHarness',(Join-Path $root 'tests/RSDirectoryHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestRSDirectory()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    @{Result=$result;Database=$scratch;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash}|ConvertTo-Json|Set-Content (Join-Path $root 'build/tests/rs-directory-results.json')
    Write-Output $result
    $depth=$app.Eval('TestSeniorApplyDepth()')
    if($depth -notlike 'PASS:*'){throw $depth}; Write-Output $depth
}finally{if($null -ne $app){$app.Quit(2)}}
