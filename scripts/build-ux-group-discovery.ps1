param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($name in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 . (Join-Path $PSScriptRoot 'apply-ux-group-discovery.ps1')
 $app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output ('PASS: compiled native group discovery: '+$OutputPath)
