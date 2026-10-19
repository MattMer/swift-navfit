param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$UseWorkingSource)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/table-sort-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($name in @($app.Forms|ForEach-Object {$_.Name})){if($name){$app.DoCmd.Close(2,$name,2)}}
 if($UseWorkingSource){foreach($m in @('modTableSort','modUXTraitGrid','modUXSeniorTable','modUXViewport')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
 $app.LoadFromText(5,'TableSortHarness',(Join-Path $root 'tests/TableSortHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestTableSort()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
