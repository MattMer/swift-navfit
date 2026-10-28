param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/pending-cells-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 $app.LoadFromText(5,'PendingCellHarness',(Join-Path $root 'tests/PendingCellHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestPendingCells()');$result;if($result -notlike 'PASS:*'){throw $result}
}finally{try{if($app.CurrentProject.AllForms.Item('frmUXComparison').IsLoaded){[void]$app.Eval('UXGridDiscardPending()')}}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
