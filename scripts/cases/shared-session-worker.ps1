param([Parameter(Mandatory=$true)][string]$App,[Parameter(Mandatory=$true)][int]$Worker,[Parameter(Mandatory=$true)][string]$Ids,[Parameter(Mandatory=$true)][string]$Start,[Parameter(Mandatory=$true)][string]$Out)
# One worker for test-shared-sessions.ps1: its own Access session on its own copy of the
# application, linked to the shared data file. Writes its outcomes to -Out.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$access=$null
try{
    $access=New-Object -ComObject Access.Application
    $access.AutomationSecurity=1;$access.OpenCurrentDatabase($App,$false)
    foreach($name in @($access.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){if($name -ne 'frmSessionKeeper'){$access.DoCmd.Close(2,$name,2)}}
    #The harness is already in this copy (the coordinator loaded it before copying):
    #loading it again would reset the session and drop its record-level hold.
    $result=[string]$access.Eval('SharedWorker('+$Worker+',"'+$Ids+'",'+$Start+')')
    [IO.File]::WriteAllText($Out,$result)
}catch{
    [IO.File]::WriteAllText($Out,'error:worker '+$Worker+': '+$_.Exception.Message)
}finally{
    if($null -ne $access){$access.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($access)}
}
