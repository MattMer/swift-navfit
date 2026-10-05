param([Parameter(Mandatory=$true)][string]$OutputPath,[Parameter(Mandatory=$true)][string]$TemplatePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($m in @('modUX','modUXSurface','modUXViewport','modUXMyWork','modUXDashboard','modUXGroupLifecycle')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
 . (Join-Path $PSScriptRoot 'apply-ux-theme.ps1')
 $app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase()
 Write-Output ('Compiled themed build: '+$OutputPath)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
