param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-block-tools-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXBlockToolsHarness',(Join-Path $root 'tests/UXBlockToolsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestUXBlockTools()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
