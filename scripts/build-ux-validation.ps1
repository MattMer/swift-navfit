param([string]$OutputPath='D:\Development\Swift NAVFIT\build\ux-r25-blank.accdb',[string]$TemplatePath='D:\Development\Swift NAVFIT\build\ux-r24-blank.accdb')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw "Refusing to overwrite $OutputPath"}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
$app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
$source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
Invoke-Expression ([regex]::Match($source,'(?s)    function Control\(.*?(?=    NewScreen ''frmSetup'')').Value)
. (Join-Path $PSScriptRoot 'apply-ux-validation.ps1')
$app.DoCmd.RunCommand(126)
$topics=Get-Content (Join-Path $root 'src/ux-help-topics.json') -Raw|ConvertFrom-Json
foreach($i in 0..8){if($app.Eval('UXHelpText('+$i+')') -cne $topics[$i].Body){throw ('Compiled help differs from topic '+$i)}}
if($app.Eval('ValidationPolicyReference("TRAIT-TWO")') -notlike '*Excerpt:*'){throw 'Validation citation did not execute'}
$app.CloseCurrentDatabase()
Write-Output "Compiled: $OutputPath"
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
