param([string]$BuildPath=(Join-Path (Split-Path $PSScriptRoot -Parent) '..\output\Swift_NAVFIT_0.25.0-dev.accdb'))
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$out=Join-Path $root 'build\tests'
[void](New-Item -ItemType Directory -Force -Path $out)
$scratch=Join-Path $out ('smoke-'+[guid]::NewGuid().ToString()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null; $db=$null; $checks=[Collections.Generic.List[string]]::new()
function Assert($condition,$message) { if(-not $condition){throw $message}; $checks.Add($message); Write-Host "PASS: $message" }
function Scalar($sql) {$r=$db.OpenRecordset($sql); try{return $r.Fields.Item(0).Value}finally{$r.Close()}}
function InvokeVba($name,[object[]]$values) {
    $literals=@($values | ForEach-Object {
        if($null -eq $_){'Null'}
        elseif($_ -is [bool]){if($_){'True'}else{'False'}}
        elseif($_ -is [datetime]){'#'+$_.ToString('yyyy-MM-dd')+'#'}
        elseif($_ -is [string]){'"'+$_.Replace('"','""')+'"'}
        else{[Convert]::ToString($_,[Globalization.CultureInfo]::InvariantCulture)}
    })
    return $app.Eval($name+'('+($literals -join ',')+')')
}
try {
    $app=New-Object -ComObject Access.Application
    Write-Host 'Created Access session'
    $app.AutomationSecurity=1 # This disposable test session only; no Trust Center settings are changed.
    $app.OpenCurrentDatabase($scratch,$true)
    Write-Host 'Opened synthetic scratch'
    $db=$app.CurrentDb()
    $app.LoadFromText(5,'SmokeHarness',(Join-Path $root 'tests\SmokeHarness.bas'))
    Assert ((Scalar 'SELECT Count(*) FROM Report') -eq 0) 'Blank report data'
    Assert ((Scalar 'SELECT Count(*) FROM Workspace') -eq 0) 'First-run setup required'
    $app.DoCmd.OpenForm('frmSetup')
    [void](InvokeVba 'SmokeSetControlValue' @('frmSetup','txtCommand','SYNTHETIC TEST COMMAND'))
    [void](InvokeVba 'SmokeSetControlValue' @('frmSetup','txtProfile','Synthetic Tester'))
    Assert ($app.Eval('SaveSetup()')) 'First-run setup action succeeds'
    Write-Host 'Workspace configured'
    $ids=@()
    foreach($kind in @('EVAL','CHIEFEVAL','FITREP')) {
        $id=InvokeVba 'CommitDraft' @(0,$kind,('Synthetic draft '+$kind),'Synthetic Member',$null,$null,"Synthetic notes`r`nSecond line",'Smoke test')
        $ids+= [int]$id
    }
    Assert ((Scalar 'SELECT Count(*) FROM Report') -eq 3) 'All three draft types persist'
    $id=$ids[0]
    $guidBefore=Scalar "SELECT ReportGuid FROM Report WHERE ReportID=$id"
    $firstRevision=Scalar "SELECT CurrentRevisionID FROM Report WHERE ReportID=$id"
    [void](InvokeVba 'CommitDraft' @($id,'EVAL','Changed title','Synthetic Member',[datetime]'2026-01-01',[datetime]'2026-09-30',"Changed notes`r`nLine two",'Edit test'))
    Assert ((Scalar "SELECT Count(*) FROM ReportRevision WHERE ReportID=$id") -eq 2) 'Edit appends a revision'
    [void](InvokeVba 'CommitDraft' @($id,'EVAL','Changed title','Synthetic Member',[datetime]'2026-01-01',[datetime]'2026-09-30',"Changed notes`r`nLine two",'No-op test'))
    Assert ((Scalar "SELECT Count(*) FROM ReportRevision WHERE ReportID=$id") -eq 2) 'No-op save does not append history'
    [void](InvokeVba 'RestoreDraft' @($id,$firstRevision))
    Assert ((Scalar "SELECT Count(*) FROM ReportRevision WHERE ReportID=$id") -eq 3) 'Restore appends instead of erasing history'
    Assert ((Scalar "SELECT Title FROM Report WHERE ReportID=$id") -eq 'Synthetic draft EVAL') 'Restore recovers earlier values'
    Assert ([Convert]::IsDBNull((Scalar "SELECT PeriodStart FROM Report WHERE ReportID=$id"))) 'Restore preserves null date'
    Assert ((Scalar "SELECT ReportGuid FROM Report WHERE ReportID=$id") -eq $guidBefore) 'Stable report identity survives restore'
    $copyID=InvokeVba 'DuplicateDraft' @($id)
    Assert ((Scalar "SELECT ReportGuid FROM Report WHERE ReportID=$copyID") -ne $guidBefore) 'Duplicate has a new stable identity'
    [void](InvokeVba 'ArchiveDraft' @($id,$true))
    $rejected=InvokeVba 'RejectArchived' @($id)
    Assert $rejected 'Archived draft edit rejected'
    Assert ((Scalar "SELECT Count(*) FROM ReportRevision WHERE ReportID=$id") -eq 3) 'Rejected edit creates no revision'
    [void](InvokeVba 'ArchiveDraft' @($id,$false))
    Assert ((Scalar "SELECT Count(*) FROM ReportEvent WHERE ReportID=$id") -eq 2) 'Archive and recovery events retained'
    $wrongType=InvokeVba 'RejectTypeChange' @($id)
    Assert $wrongType 'Existing report type change rejected'
    Assert (InvokeVba 'RejectSnapshotFailure' @($id)) 'Injected snapshot write failure rejected'
    Assert ((Scalar "SELECT Title FROM Report WHERE ReportID=$id") -eq 'Synthetic draft EVAL') 'Failed transaction rolls back report values'
    Assert ((Scalar "SELECT Count(*) FROM ReportRevision WHERE ReportID=$id") -eq 3) 'Failed transaction rolls back revision history'
    Assert ((Scalar "SELECT Count(*) FROM RevisionValue WHERE RevisionID NOT IN (SELECT RevisionID FROM ReportRevision)") -eq 0) 'No orphan snapshots after rollback'
    $blankID=InvokeVba 'CommitDraft' @(0,'EVAL','','',$null,$null,'','Incomplete draft test')
    Assert ($blankID -gt 0) 'Incomplete drafts with empty fields save'
    $rootFolder=InvokeVba 'CreateFolder' @('Synthetic Command',0)
    $childFolder=InvokeVba 'CreateFolder' @('Division A',$rootFolder)
    Assert ((InvokeVba 'FolderPath' @($childFolder)) -eq 'Synthetic Command / Division A') 'Folder hierarchy path resolves'
    Assert (InvokeVba 'RejectFolderCycle' @($rootFolder,$childFolder)) 'Folder cycle rejected'
    $revisionBeforeMove=Scalar "SELECT CurrentRevisionID FROM Report WHERE ReportID=$id"
    [void](InvokeVba 'AssignFolder' @($id,$childFolder))
    Assert ((Scalar "SELECT FolderID FROM Report WHERE ReportID=$id") -eq $childFolder) 'Report moves to destination folder'
    Assert ((Scalar "SELECT CurrentRevisionID FROM Report WHERE ReportID=$id") -eq $revisionBeforeMove) 'Folder move preserves committed report revision'
    Assert ((Scalar 'SELECT Count(*) FROM FolderMove') -eq 1) 'Folder move has contributor history'
    [void](InvokeVba 'MoveFolder' @($childFolder,0))
    Assert ((InvokeVba 'FolderPath' @($childFolder)) -eq 'Division A') 'Folder can move to root'
    foreach($name in @('frmHome','frmLibrary','frmEditor','frmSupport')){
        if($name -eq 'frmEditor'){$app.DoCmd.OpenForm($name,0,'','',0,0,[string]$id)}else{$app.DoCmd.OpenForm($name)}
        Assert ($app.Eval('Forms("'+$name+'").Controls.Count') -gt 0) "Form opens: $name"
        if($name -eq 'frmLibrary'){Assert ($app.Eval('(Forms!frmLibrary!lstReports.ListCount - Abs(Forms!frmLibrary!lstReports.ColumnHeads)) = 5')) 'Library initially displays active reports'}
        if($name -eq 'frmEditor'){[void]$app.Eval('UXSelectSection(9)');Assert ($app.Eval('(Forms!frmEditor!lstRevisions.ListCount - Abs(Forms!frmEditor!lstRevisions.ColumnHeads)) = 3')) 'History section displays all revisions'}
        if(-not $app.Eval('SmokeFormFits("'+$name+'")')){throw "Control outside form width: $name"}
        $app.DoCmd.Close(2,$name,2)
    }
    $app.SaveAsText(2,'frmEditor',(Join-Path $out 'editor-form.txt'))
    Assert $true 'Visible controls stay within form width'
    $app.DoCmd.OpenForm('frmLibrary')
    $smokeGroup=InvokeVba 'CreateSummaryGroup' @('Smoke group','EVAL','E5')
    [void](InvokeVba 'SmokeAddGroupMember' @($smokeGroup,$id,'E5',''))
    [void]$app.Eval('UXRefreshGroupFilter()')
    [void](InvokeVba 'SmokeSetControlValue' @('frmLibrary','lstFolders',$smokeGroup))
    [void]$app.Eval('RefreshLibrary()')
    Assert ($app.Eval('(Forms!frmLibrary!lstReports.ListCount - Abs(Forms!frmLibrary!lstReports.ColumnHeads)) = 1')) 'Summary-group selection filters reports'
    [void](InvokeVba 'SmokeSetControlValue' @('frmLibrary','lstReports',$id))
    Assert ($app.Eval('UXChooseReportGroup()')) 'Move report opens summary-group selector'
    [void](InvokeVba 'SmokeSetControlValue' @('frmUXMoveGroup','cboGroup',0))
    Assert ($app.Eval('UXPreviewMoveGroup()')) 'Summary-group move preview succeeds'
    Assert ($app.Eval('UXApplyMoveGroup()')) 'Summary-group move apply succeeds'
    $app.DoCmd.Close(2,'frmUXMoveGroup',2)
    Assert ($app.Eval('Forms!frmLibrary!lstReports.ListCount = 0')) 'Move refreshes current summary-group view'
    [void](InvokeVba 'SmokeAddGroupMember' @($smokeGroup,$id,'E5',''))
    [void](InvokeVba 'SmokeSetControlValue' @('frmLibrary','lstFolders',-1))
    [void]$app.Eval('RefreshLibrary()')
    Assert ($app.Eval('(Forms!frmLibrary!lstReports.ListCount - Abs(Forms!frmLibrary!lstReports.ColumnHeads)) = 4')) 'Ungrouped view locates drafts outside summary groups'
    [void](InvokeVba 'SmokeSetControlValue' @('frmLibrary','lstFolders',0))
    [void](InvokeVba 'SmokeSetControlValue' @('frmLibrary','txtSearch',"O'Brien"))
    [void]$app.Eval('RefreshLibrary()')
    Assert ($app.Eval('Forms!frmLibrary!lstReports.ListCount = 0')) 'Search safely handles apostrophes'
    [void](InvokeVba 'SmokeSetControlValue' @('frmLibrary','txtSearch','Synthetic draft EVAL'))
    [void]$app.Eval('RefreshLibrary()')
    # Word matching also finds EVAL within CHIEFEVAL, alongside the EVAL and its copy.
    Assert ($app.Eval('(Forms!frmLibrary!lstReports.ListCount - Abs(Forms!frmLibrary!lstReports.ColumnHeads)) = 3')) 'Word search finds EVAL and CHIEFEVAL title matches'
    $app.DoCmd.Close(2,'frmLibrary',2)
    $app.DoCmd.OpenForm('frmEditor',0,'','',0,0,'0')
    [void]$app.Eval('SmokeEditorEntries()')
    Assert ($app.Eval('SaveEditor()')) 'Editor save action succeeds'
    Assert ($app.Eval('Left(Forms!frmEditor!lblSaved.Caption,6)="Saved "')) 'Editor displays save confirmation'
    [void]$app.Eval('ShowHelp("Saving")')
    Assert ($app.Eval('Screen.ActiveForm.Name') -eq 'frmUXHelp') 'Help becomes the active form'
    Assert ($app.Eval('InStr(Forms!frmUXHelp!txtGuide,"Save draft creates an atomic revision") > 0')) 'Help contains topic-specific guidance'
    [void]$app.Eval('CloseHelp()')
    Assert ($app.Eval('Screen.ActiveForm.Name') -eq 'frmEditor') 'Closing help returns to the editor'
    $app.DoCmd.Close(2,'frmEditor',2)
    $db.Close();$db=$null
    $app.CloseCurrentDatabase()
    $app.OpenCurrentDatabase($scratch,$true)
    $db=$app.CurrentDb()
    Assert ((Scalar 'SELECT Count(*) FROM Report') -eq 6) 'Reports survive close and reopen'
    Assert ((Scalar 'SELECT Count(*) FROM ReportRevision') -eq 8) 'Revision history survives close and reopen'
    $result=@{Status='PASS';AccessVersion=$app.Version;AccessBuild=$app.Build;TestDatabase=$scratch;Checks=$checks.ToArray();BuildPath=[IO.Path]::GetFullPath($BuildPath);BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;NAVFITCompatibility='Not tested';Runtime='Not tested'}
    $result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $out 'smoke-results.json') -Encoding utf8
    $result|ConvertTo-Json -Depth 5
} finally {
    if($null -ne $db){$db.Close()}
    if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
