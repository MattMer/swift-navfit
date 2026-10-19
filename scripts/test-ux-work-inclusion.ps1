param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/work-inclusion-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($pair in @(@('modUXReadRows','UXPrepareWorkRows'),@('modUXMyWork','UXWorkCount'))){
  $module=Join-Path $root ('build/tests/'+$pair[0]+'-'+[guid]::NewGuid()+'.bas')
  $source=[IO.File]::ReadAllText((Join-Path $root ('src/'+$pair[0]+'.bas')))
  $legacy=[IO.File]::ReadAllText((Join-Path $root ('tests/'+$pair[1]+'Legacy.txt')))
  [IO.File]::WriteAllText($module,$source+"`r`n"+$legacy);$app.LoadFromText(5,$pair[0],$module)
 }
 $app.LoadFromText(5,'WorkInclusionHarness',(Join-Path $root 'tests/WorkInclusionHarness.bas'));$app.DoCmd.RunCommand(126)
 $out=Join-Path $root 'build/batch5-work-inclusion-scale.csv'
 $result=$app.Eval('TestWorkInclusion("'+$out+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $source=[IO.File]::ReadAllText((Join-Path $root 'src/modUXMyWork.bas'))
 $source=$source.Replace('Public Function UXWorkVisible(ByVal kind As String, ByVal id As Long, ByVal pid As Long) As Boolean','Public Function UXWorkVisible(ByVal kind As String, ByVal id As Long, ByVal pid As Long) As Boolean'+"`r`n"+'    Err.Raise vbObjectError + 8793, , "Single-item visibility callback invoked"')
 $module=Join-Path $root ('build/tests/no-work-callback-'+[guid]::NewGuid()+'.bas')
 [IO.File]::WriteAllText($module,$source+"`r`n"+[IO.File]::ReadAllText((Join-Path $root 'tests/UXWorkCountLegacy.txt')));$app.LoadFromText(5,'modUXMyWork',$module);$app.DoCmd.RunCommand(126)
 if(-not $app.Eval('TestVisibilityStub()')){throw 'Legacy visibility rejection probe not installed'}
 Write-Output 'PASS: single-item visibility callback rejection probe is active'
 $result=$app.Eval('TestNoVisibilityCallbacks()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
