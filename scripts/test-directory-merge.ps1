param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/directory-merge-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'DirectoryMergeHarness',(Join-Path $root 'tests/DirectoryMergeHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestDirectoryMerge()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 $app.CloseCurrentDatabase()
 $app.OpenCurrentDatabase($scratch,$true);$app.DoCmd.DeleteObject(5,'DirectoryMergeHarness');$app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase()
 @{Result=$result;Fixture=$scratch;SHA256=(Get-FileHash $scratch).Hash;Reopened=$true}|ConvertTo-Json|Set-Content (Join-Path $root 'build/directory-merge-verification.json')
 Write-Output ('Merge-history migration fixture: '+$scratch)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
