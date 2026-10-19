param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath,[string]$SeedWorkspacePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 foreach($m in @('modReportPages','modReportPageMap','modUXHelp')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
 if($SeedWorkspacePath){
  $source=[IO.Path]::GetFullPath($SeedWorkspacePath);$before=(Get-FileHash -LiteralPath $source).Hash
  $result=$app.Eval('MigrateWorkspace("'+$source.Replace('"','""')+'")')
  if($result -notlike 'PASS:*'){throw $result}
  if((Get-FileHash -LiteralPath $source).Hash -ne $before){throw 'Seed source changed'}
 }
 $app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output ('PASS: report rail and front-page traits build: '+$OutputPath)
