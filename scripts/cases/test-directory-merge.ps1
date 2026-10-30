param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$DisplayFaultOnly)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/directory-merge-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($DisplayFaultOnly){
  $code=[IO.File]::ReadAllText((Join-Path $root 'src/modUXDirectoryMerge.bas'))
  $anchor='    id = ApplyDirectoryMerge(Nz(f!cboSource, 0), Nz(f!cboTarget, 0), reviewedStamp, Nz(f!txtReason, ""))'
  if(([regex]::Matches($code,[regex]::Escape($anchor))).Count -ne 1){throw 'Expected merge display fault boundary missing or duplicated'}
  $instrumented=$scratch+'.bas';[IO.File]::WriteAllText($instrumented,$code.Replace($anchor,$anchor+"`r`n"+'    Err.Raise 5, "Synthetic merge display", "SYNTHETIC merge display failure"'));$app.LoadFromText(5,'modUXDirectoryMerge',$instrumented)
 }
 $app.LoadFromText(5,'DirectoryMergeHarness',(Join-Path $root 'tests/DirectoryMergeHarness.bas'));$app.DoCmd.RunCommand(126)
 $expression=if($DisplayFaultOnly){'TestMergeDisplayFailure()'}else{'TestDirectoryMerge()'}
 $result=$app.Eval($expression);Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 $app.CloseCurrentDatabase()
 $app.OpenCurrentDatabase($scratch,$true);$app.DoCmd.DeleteObject(5,'DirectoryMergeHarness');$app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase()
 @{Result=$result;Fixture=$scratch;SHA256=(Get-FileHash $scratch).Hash;Reopened=$true;DisplayFaultOnly=[bool]$DisplayFaultOnly}|ConvertTo-Json|Set-Content (Join-Path $root 'build/directory-merge-verification.json')
 Write-Output ('Merge-history migration fixture: '+$scratch)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
