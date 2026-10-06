param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$BuildPath=[IO.Path]::GetFullPath($BuildPath)
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($BuildPath,$true)
 $app.LoadFromText(5,'DirectoryDemoSeed',(Join-Path $root 'tests/DirectoryDemoSeed.bas'))
 $result=$app.Eval('SeedDirectoryDemoHistory()')
 if($result -notlike 'PASS:*'){throw $result}
 $app.DoCmd.DeleteObject(5,'DirectoryDemoSeed');$app.DoCmd.RunCommand(126)
 $app.CloseCurrentDatabase();Write-Output $result
} finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
