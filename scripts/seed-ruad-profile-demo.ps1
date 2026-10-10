param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent;$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase([IO.Path]::GetFullPath($BuildPath),$true)
 if($app.Eval('DCount("*","CommandMember","SourceKind=''Demo RUAD''")') -ne 0){throw 'Demo profiles already seeded'}
 $app.LoadFromText(5,'UXRUADSeed',(Join-Path $root 'tests/UXRUADSeed.bas'))
 $count=$app.Eval('SeedRUADDemo()');if($count -ne 2){throw 'Expected two synthetic RUAD profiles'}
 $app.DoCmd.DeleteObject(5,'UXRUADSeed');$app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase()
 Write-Output 'Added two clearly labeled synthetic RUAD profiles; existing member records were not edited.'
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
