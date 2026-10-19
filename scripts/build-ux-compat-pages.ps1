param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath,[switch]$FromR110)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 for($wait=0;$wait -lt 30 -and $app.Forms.Count -eq 0;$wait++){Start-Sleep -Milliseconds 100}
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 if($FromR110){
  foreach($patch in @('group-discovery','member-psg','import-rank','native-dates')){. (Join-Path $PSScriptRoot ('apply-ux-'+$patch+'.ps1'))}
 }
 Write-Output 'Applying fixed pages and workspace import'
 . (Join-Path $PSScriptRoot 'apply-ux-compat-pages.ps1')
 Write-Output 'Applying table sorting'
 . (Join-Path $PSScriptRoot 'apply-ux-table-sort.ps1')
 Write-Output 'Applying existing application palette'
 . (Join-Path $PSScriptRoot 'apply-ux-theme.ps1')
 $app.DoCmd.RunCommand(126)
}catch{
 Write-Output $_.ScriptStackTrace
 throw
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output ('PASS: compiled compatible workspace import and fixed report pages: '+$OutputPath)
