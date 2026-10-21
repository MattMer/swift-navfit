if(-not $script:swiftManagedBuild){foreach($m in @('modUXReadRows','modUXGroupOrganization','modWorkspace','modUX','modGroupPolicy','modUXGroupLifecycle')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary')
[void](Control $f 100 'uxGroupOrganization' 0 0 100 100 'Groups');$script:swiftControl.Visible=$false
$f.Controls('cboDestination').Visible=$false
$f.Controls('cmdMoveFolder').Visible=$false
$f.Controls('cmdMoveReport').OnClick='=UXChooseReportGroup()';$f.Controls('cmdMoveReport').Caption='Move to summary group'
$f.Controls('uxFolderFilter').ControlTipText='Filter reports by summary group; Ungrouped shows reports with no group.'
SaveScreen $f 'frmLibrary'
NewScreen 'frmUXMoveGroup' 'Move report to summary group' 8200;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.Width=10500
[void](Control $f 109 'reportID' 0 0 100 100);$script:swiftControl.Visible=$false
[void](Control $f 100 'lblMember' 400 1000 9500 400 '')
[void](Control $f 111 'cboGroup' 400 1500 9500 450);$c=$script:swiftControl;$c.RowSourceType='Table/Query';$c.ColumnCount=2;$c.ColumnWidths='0;9000';$c.LimitToList=$true
[void](Control $f 104 'cmdPreview' 400 2150 4500 500 'Preview move' '=UXPreviewMoveGroup()')
[void](Control $f 104 'cmdApply' 5400 2150 4500 500 'Apply previewed move' '=UXApplyMoveGroup()')
[void](Control $f 109 'txtPreview' 400 2900 9500 3300);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2
[void](Control $f 100 'lblStatus' 400 6400 9500 700 '')
[void](Control $f 104 'cmdClose' 400 7300 9500 500 'Back to reports' '=UXCloseScreen("frmUXMoveGroup")')
$f.OnLoad='=UXLoadMoveGroup()';SaveScreen $f 'frmUXMoveGroup'
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups')
[void](Control $f 109 'uxStateStamp' 0 0 100 100);$script:swiftControl.Visible=$false
[void](Control $f 100 'uxStateLabel' 2850 8650 4500 350 'Local group status (does not change report review)')
[void](Control $f 111 'uxGroupState' 2850 9100 6600 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='Initialized;In-Work;Ready for Signature;Completed';$c.LimitToList=$true
[void](Control $f 104 'uxSaveState' 9750 9100 3550 450 'Set group status' '=UXSetGroupWorkflow()')
$f.Section(0).Height=9950;$f.ScrollBars=2
SaveScreen $f 'frmUXGroups'
