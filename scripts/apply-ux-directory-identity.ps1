$app.LoadFromText(5,'modDirectoryBatch',(Join-Path $root 'src/modDirectoryBatch.bas'))
$app.LoadFromText(5,'modUXDirectory',(Join-Path $root 'src/modUXDirectory.bas'))
$app.DoCmd.OpenForm('frmUXDirectoryBatch',1);$f=$app.Forms.Item('frmUXDirectoryBatch')
[void](Control $f 100 'lblFrom' 7000 1000 2800 350 'Report from date');[void](Control $f 100 'lblTo' 10200 1000 2800 350 'Report to date');SaveScreen $f 'frmUXDirectoryBatch'
$app.DoCmd.OpenForm('frmUXMemberHistory',1);$f=$app.Forms.Item('frmUXMemberHistory')
[void](Control $f 104 'cmdConfirm' 400 7150 12600 500 'Confirm member identity for selected Swift report' '=UXHistoryConfirmIdentity()');$f.Section(0).Height=8200;SaveScreen $f 'frmUXMemberHistory'
NewScreen 'frmUXConfirmIdentity' 'Confirm saved report member identity' 7200;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.ScrollBars=2
[void](Control $f 109 'txtPreview' 400 1000 12600 2400);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
foreach($n in @('txtReport','txtRevision','txtMember','txtVersion')){[void](Control $f 109 $n 0 0 1 1);$script:swiftControl.Visible=$false}
[void](Control $f 100 'lblReason' 400 3650 12600 350 'Reason for confirming this is the same person')
[void](Control $f 109 'txtReason' 400 4150 12600 600)
[void](Control $f 100 'lblStatus' 400 5000 12600 600 '')
[void](Control $f 104 'cmdApply' 400 6100 6100 500 'Confirm identity link' '=UXHistoryApplyIdentity()')
[void](Control $f 104 'cmdCancel' 6800 6100 6200 500 'Cancel' '=UXDirectoryClose("frmUXConfirmIdentity")')
SaveScreen $f 'frmUXConfirmIdentity'
