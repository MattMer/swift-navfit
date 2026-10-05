param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-workflows-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXWorkflowHarness',(Join-Path $root 'tests/UXWorkflowHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestUXWorkflows()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
foreach($invalid in @($false,$true)){
    $quitScratch=Join-Path $root ('build/tests/ux-quit-'+[guid]::NewGuid()+'.accdb')
    Copy-Item -LiteralPath $BuildPath -Destination $quitScratch
    $quitApp=New-Object -ComObject Access.Application
    try{
        $quitApp.AutomationSecurity=1;$quitApp.OpenCurrentDatabase($quitScratch,$true)
        $quitApp.LoadFromText(5,'UXWorkflowHarness',(Join-Path $root 'tests/UXWorkflowHarness.bas'))
        [void]$quitApp.Eval('PrepareForcedQuit('+$(if($invalid){'True'}else{'False'})+')')
        $quitApp.Quit(2)
    }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($quitApp)}
    $reader=New-Object -ComObject DAO.DBEngine.120
    $db=$reader.OpenDatabase($quitScratch,$true,$true)
    try{
        $saved=$db.OpenRecordset('SELECT OfficialName,MemberSSN FROM Report',4)
        $pending=$db.OpenRecordset('SELECT OfficialName,MemberSSN FROM RecoveryDraft',4)
        if($invalid){
            if($saved.Fields.Item('MemberSSN').Value -isnot [DBNull] -or $pending.EOF -or $pending.Fields.Item('MemberSSN').Value -cne '123' -or $pending.Fields.Item('OfficialName').Value -cne 'QUIT SAVE NAME'){throw 'Forced quit lost invalid pending values or committed them.'}
        }else{
            if($saved.Fields.Item('OfficialName').Value -cne 'QUIT SAVE NAME' -or $saved.Fields.Item('MemberSSN').Value -cne '000000777' -or -not $pending.EOF){throw 'Forced quit did not save valid edits.'}
        }
        $saved.Close();$pending.Close()
    }finally{$db.Close()}
}
Write-Output 'PASS: forced Access shutdown saves valid edits and retains invalid raw edits for recovery'
