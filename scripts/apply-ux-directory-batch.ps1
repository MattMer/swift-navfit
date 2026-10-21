if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modDirectoryBatch',(Join-Path $root 'src/modDirectoryBatch.bas'))}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXDirectory',(Join-Path $root 'src/modUXDirectory.bas'))}
$app.DoCmd.OpenForm('frmUXDirectory',1);$f=$app.Forms.Item('frmUXDirectory')
[void](Control $f 104 'cmdCreateReports' 400 7300 12600 500 'Create group reports from directory members' '=UXDirectoryBatchOpen()');$f.Section(0).Height=8300;SaveScreen $f 'frmUXDirectory'
NewScreen 'frmUXDirectoryBatch' 'Create reports from command members' 11200;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblChoose' 400 1000 12600 350 'Choose summary group and report period')
[void](Control $f 111 'cboGroup' 400 1450 6300 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;6000';$c.LimitToList=$true
[void](Control $f 109 'txtFrom' 7000 1450 2800 450);$script:swiftControl.ControlTipText='Report from date'
[void](Control $f 109 'txtTo' 10200 1450 2800 450);$script:swiftControl.ControlTipText='Report to date'
[void](Control $f 110 'lstMembers' 400 2250 12600 2300);$c=$script:swiftControl;$c.MultiSelect=1;$c.ColumnCount=3;$c.ColumnWidths='0;9400;2800'
[void](Control $f 104 'cmdPreview' 400 4850 12600 500 'Preview new reports' '=UXDirectoryBatchPreview()')
[void](Control $f 109 'txtPreview' 400 5600 12600 2600);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 109 'txtStamp' 0 0 1 1);$script:swiftControl.Visible=$false
[void](Control $f 100 'lblStatus' 400 8500 12600 650 '')
[void](Control $f 104 'cmdApply' 400 9600 6100 500 'Create previewed drafts' '=UXDirectoryBatchApply()');$script:swiftControl.Enabled=$false
[void](Control $f 104 'cmdClose' 6800 9600 6200 500 'Back to directory' '=UXDirectoryClose("frmUXDirectoryBatch")')
SaveScreen $f 'frmUXDirectoryBatch'
