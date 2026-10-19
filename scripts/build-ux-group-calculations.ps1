param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
$before=(Get-FileHash -LiteralPath $TemplatePath).Hash
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 foreach($m in @('modUXDashboard','modRSDirectory','modRSDirectoryUI','modUXCalculationContext','modUXResourceTools','modUXGroupPages')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
 $app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash -LiteralPath $TemplatePath).Hash -ne $before){throw 'Template changed'}
Write-Output ('PASS: compiled group-calculation candidate: '+$OutputPath)
