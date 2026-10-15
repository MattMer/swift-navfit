param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($form in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$form,2)}
 foreach($name in @('modNAVFITUpgrade','modImportReview','modUXNativeGroups','modNAVFITExchange')){$app.LoadFromText(5,$name,(Join-Path $root ('src/'+$name+'.bas')))}
 $app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output ('PASS: compiled automatic v33 import upgrade: '+$OutputPath)
