param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/work-inclusion-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($moduleName in @('modUXReadRows','modUXMyWork')){
  $module=Join-Path $root ('build/tests/'+$moduleName+'-'+[guid]::NewGuid()+'.bas')
  $source=[IO.File]::ReadAllText((Join-Path $root ('src/'+$moduleName+'.bas')))
  [IO.File]::WriteAllText($module,$source);$app.LoadFromText(5,$moduleName,$module)
 }
 $app.LoadFromText(5,'WorkInclusionHarness',(Join-Path $root 'tests/WorkInclusionHarness.bas'));$app.DoCmd.RunCommand(126)
 $out=Join-Path $root 'build/batch5-work-inclusion-scale.csv'
 $result=$app.Eval('TestWorkInclusion("'+$out+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $source=[IO.File]::ReadAllText((Join-Path $root 'src/modUXMyWork.bas'))
 $source=$source.Replace('Public Function UXWorkVisible(ByVal kind As String, ByVal id As Long, ByVal pid As Long) As Boolean','Public Function UXWorkVisible(ByVal kind As String, ByVal id As Long, ByVal pid As Long) As Boolean'+"`r`n"+'    Err.Raise vbObjectError + 8793, , "Single-item visibility callback invoked"')
 $module=Join-Path $root ('build/tests/no-work-callback-'+[guid]::NewGuid()+'.bas')
 [IO.File]::WriteAllText($module,$source);$app.LoadFromText(5,'modUXMyWork',$module);$app.DoCmd.RunCommand(126)
 if(-not $app.Eval('TestVisibilityStub()')){throw 'Legacy visibility rejection probe not installed'}
 Write-Output 'PASS: single-item visibility callback rejection probe is active'
 $result=$app.Eval('TestNoVisibilityCallbacks()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
