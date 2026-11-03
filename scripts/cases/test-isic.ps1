param([Parameter(Mandatory=$true)][string]$BuildPath)
#ISICs for summary letters, the reporting senior e-mail and the letter's
#lower section (ISIC address and UIC, senior address, e-mail, phone, on time).
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/isic-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ISICHarness',(Join-Path $root 'tests/ISICHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestISIC()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
