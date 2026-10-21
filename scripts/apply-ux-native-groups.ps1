if(-not $script:swiftManagedBuild){foreach($m in @('modImportReview','modUXNativeGroups','modOfficialBatch')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
NewScreen 'frmUXNativeGroups' 'Import NAVFIT groups' 13500;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 109 'txtSource' 400 950 12600 400);$script:swiftControl.Locked=$true
[void](Control $f 100 'lblCentury' 400 1500 9500 400 'Choose counseling date century before loading and previewing groups')
[void](Control $f 111 'cboCentury' 10300 1500 2700 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='1900;2000';$c.LimitToList=$true;$c.AfterUpdate='=UXLoadNativeGroups()'
[void](Control $f 110 'lstGroups' 400 2150 12600 2350);$c=$script:swiftControl;$c.ColumnCount=5;$c.ColumnWidths='0;7000;2100;1500;1000';$c.RowSourceType='Value List';$c.MultiSelect=1;$c.OnClick='=UXNativeGroupSelectionChanged()'
[void](Control $f 104 'cmdPreview' 400 4900 6100 450 'Preview selected groups and reports' '=UXPreviewNativeGroups()')
[void](Control $f 104 'cmdApply' 6800 4900 6200 450 'Apply selected groups' '=UXApplyNativeGroups()');$script:swiftControl.Enabled=$false
[void](Control $f 109 'txtPreview' 400 5700 12600 4800);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2
[void](Control $f 100 'lblStatus' 400 10800 12600 1050 'Choose the counseling century, then select groups and Preview.')
[void](Control $f 104 'cmdBack' 400 12200 12600 500 'Back to import choices' '=UXCloseNativeGroups()')
SaveScreen $f 'frmUXNativeGroups'
$app.DoCmd.OpenForm('frmUXImportChoose',1);$f=$app.Forms.Item('frmUXImportChoose');$f.Controls('cmdNAVFIT').Caption='Choose NAVFIT database - groups and reports';$f.Controls('cmdNAVFIT').OnClick='=UXChooseNativeGroups()';SaveScreen $f 'frmUXImportChoose'
