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
 . (Join-Path $PSScriptRoot 'apply-report-tools.ps1')
 $app.DoCmd.RunCommand(126)
 [void]$app.Eval('RSLatestBaseline(0,"E6")')
}finally{
 $app.Quit(2)
 foreach($variable in @(Get-Variable -Scope Script)){
  if($variable.Name -eq 'app'){continue}
  if($null -ne $variable.Value -and [Runtime.InteropServices.Marshal]::IsComObject($variable.Value)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($variable.Value)}catch{}}
 }
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
 [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
}
if((Get-FileHash -LiteralPath $TemplatePath).Hash -ne $before){throw 'Template changed'}
Write-Output ('PASS: compiled embedded report tools candidate: '+$OutputPath)
