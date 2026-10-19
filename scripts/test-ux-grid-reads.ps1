param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/grid-reads-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$module=Join-Path $root 'build/grid-reads-test-module.bas'
$source=Get-Content (Join-Path $root 'src/modUXTraitGrid.bas') -Raw
$legacy=Get-Content (Join-Path $root 'tests/UXGridLoadLegacy.txt') -Raw
[IO.File]::WriteAllText($module,$source+"`r`n"+$legacy)
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 $app.LoadFromText(5,'modUXTraitGrid',$module)
 $app.LoadFromText(5,'GridReadHarness',(Join-Path $root 'tests/GridReadHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $out=Join-Path $root 'build/batch4-grid-scale.csv'
 $result=$app.Eval('TestGridReads("'+$out+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
}finally{try{if($app.CurrentProject.AllForms.Item('frmUXComparison').IsLoaded){[void]$app.Eval('UXGridDiscardPending()')}}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
