param([Parameter(Mandatory=$true)][string]$WorkspacePath,[Parameter(Mandatory=$true)][string]$ReturnedPath,[Parameter(Mandatory=$true)][int]$ReportID,[string]$Marker='',[switch]$AlreadyApplied)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($WorkspacePath,$true)
 $app.LoadFromText(5,'NAVFITNativeReturnHarness',(Join-Path $root 'tests/NAVFITNativeReturnHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('VerifyNativeReturn('+ $ReportID +',"'+$ReturnedPath+'","'+(Join-Path $root 'NAVFIT_BLANK.accdb')+'","'+$Marker+'",'+$(if($AlreadyApplied){'True'}else{'False'})+')')
 if($result -notlike 'PASS:*'){throw $result}
 $app.CloseCurrentDatabase()
 $reexport=$result.Split(@('; reexport='),[StringSplitOptions]::None)[1]
 @{Status=$result;Workspace=$WorkspacePath;ReturnedFile=$ReturnedPath;ReturnedSHA256=(Get-FileHash -LiteralPath $ReturnedPath).Hash;Reexport=$reexport;ReexportSHA256=(Get-FileHash -LiteralPath $reexport).Hash}|ConvertTo-Json|Set-Content -LiteralPath ($ReturnedPath+'.native-proof.json')
 Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
