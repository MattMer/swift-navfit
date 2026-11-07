param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$folder=Join-Path $root ('build/tests/contributor-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $folder)
$source=Join-Path $folder 'synthetic.accdb';$target=Join-Path $folder 'imported.accdb';$cancel=Join-Path $folder 'cancel-blank.accdb'
foreach($path in @($source,$target,$cancel)){Copy-Item -LiteralPath $BuildPath -Destination $path}
$app=New-Object -ComObject Access.Application
function AssertChooser {
 if(-not $app.Eval('CurrentProject.AllForms("frmContributorStart").IsLoaded') -or $app.Eval('HasActiveContributor()')){throw 'Initialized launch omitted explicit contributor choice'}
 if($null -ne $app.Eval('Forms!frmContributorStart!cboContributor.Value') -and $app.Eval('Forms!frmContributorStart!cboContributor.Value') -isnot [DBNull]){throw 'Contributor silently preselected'}
}
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
 $app.LoadFromText(5,'ContributorStartupHarness',(Join-Path $root 'tests/ContributorStartupHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestContributorStartup()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $safety=$app.Eval('TestWorkspaceSafety()');Write-Output $safety
 if($safety -notlike 'PASS:*'){throw $safety}
 $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($source,$true)
 AssertChooser
 if($app.Eval('Forms!frmContributorStart!cboContributor.ListCount') -ne 3){throw 'Reopened contributor list lost rows'}
 $app.CloseCurrentDatabase();$sourceHash=(Get-FileHash -LiteralPath $source).Hash
 $app.OpenCurrentDatabase($target,$true)
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC IMPORT BOOTSTRAP","Temporary bootstrap")')
 if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")') -or -not $app.Eval('UXApplySwiftWorkspace()')){throw 'Settings & Help import failed'}
 if($app.Eval('HasActiveContributor()')){throw 'Imported workspace retained confirmed session'}
 [void]$app.Eval('SwiftAppStartup()');AssertChooser
 if($app.Eval('Forms!frmContributorStart!cboContributor.ListCount') -ne 3){throw 'Imported contributor list incomplete'}
 $app.Forms.Item('frmContributorStart').Controls.Item('cboContributor').Value= [int]$app.Eval('DMin("UserProfileID","UserProfile")')
 if(-not $app.Eval('AcceptContributorStartup()')){throw 'Imported contributor choice failed'}
 $author=[int]$app.Eval('ActiveContributorID()');$report=[int]$app.Eval('DMin("ReportID","Report")')
 $message=[int]$app.Eval('AddReportMessage('+$report+',"Explicit imported contributor")')
 if($app.Eval('DLookup("AuthorProfileID","ReportMessage","MessageID='+$message+'")') -ne $author){throw 'Imported comment actor wrong'}
 $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($target,$true);AssertChooser
 # No pending editor exists here. Exercise the real Exit callback, which quits Access.
 try{[void]$app.Eval('CancelContributorStartup()')}catch{if($_.Exception.Message -notmatch 'disconnected|RPC|remote procedure|COM|0x800'){throw}}
}finally{try{$app.Quit(2)}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Import altered source workspace'}
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($cancel,$true)
 if(-not $app.Eval('CurrentProject.AllForms("frmSetup").IsLoaded')){throw 'Blank cancel omitted setup'}
 try{[void]$app.Eval('CancelContributorStartup()')}catch{if($_.Exception.Message -notmatch 'disconnected|RPC|remote procedure|COM|0x800'){throw}}
}finally{try{$app.Quit(2)}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
$dao=New-Object -ComObject DAO.DBEngine.120
try{
 $db=$dao.OpenDatabase($cancel,$false,$true)
 try{foreach($table in @('Workspace','UserProfile','Report','ReportRevision')){$rows=$db.OpenRecordset('SELECT Count(*) FROM ['+$table+']',4);try{if($rows.Fields.Item(0).Value -ne 0){throw ('Blank cancellation wrote '+$table)}}finally{$rows.Close()}}}finally{$db.Close()}
 $db=$dao.OpenDatabase($target,$false,$true)
 try{$rows=$db.OpenRecordset('SELECT ActiveProfileID FROM Workspace',4);try{if($rows.Fields.Item(0).Value -ne $author){throw 'Cancel changed retained contributor'}}finally{$rows.Close()}}finally{$db.Close()}
}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao)}
# Two real sessions on one shared file: the second launch warns before anyone works.
$shared=Join-Path $folder 'shared.accdb';Copy-Item -LiteralPath $source -Destination $shared
$first=$null;$second=$null
try{
 $first=New-Object -ComObject Access.Application;$first.AutomationSecurity=1;$first.OpenCurrentDatabase($shared,$false)
 if($first.Eval('Forms!frmContributorStart!uxWorkspaceWarning.Visible')){throw ('Lone session warned: '+$first.Eval('Forms!frmContributorStart!uxWorkspaceWarning.Caption'))}
 $second=New-Object -ComObject Access.Application;$second.AutomationSecurity=1;$second.OpenCurrentDatabase($shared,$false)
 if($second.CurrentProject.FullName -ne $shared){throw 'Second shared session did not open'}
 $warning=[string]$second.Eval('Forms!frmContributorStart!uxWorkspaceWarning.Caption')
 if(-not $second.Eval('Forms!frmContributorStart!uxWorkspaceWarning.Visible') -or $warning -notlike '*already open in another session (this computer)*'){throw ('Second session not warned: '+$warning)}
 $second.CloseCurrentDatabase()
 Write-Output ('PASS: second shared session warned: '+$warning)
}finally{
 if($second){try{$second.Quit(2)}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($second)}
 if($first){try{$first.CloseCurrentDatabase()}catch{};try{$first.Quit(2)}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($first)}
}
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){throw 'Test changed supplied build'}
@{Status='PASS';BuildPath=$BuildPath;BuildSHA256=$hash;SourceSHA256=$sourceHash;Assertions=$result;Reopen=$true;SettingsImport=$true;ImportedAttribution=$true;BlankCancel=$true;InitializedCancel=$true;SourceUnchanged=$true}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $folder 'verification.json')
Write-Output ('PASS: reopen, Settings & Help import, imported attribution and blank/initialized cancellation; '+$folder)
