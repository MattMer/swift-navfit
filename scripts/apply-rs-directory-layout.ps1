# Invoked in the shared Access build scope, after the UX layout helpers.
$rsKeys=@($uxFields|Where-Object {$_.key.StartsWith('ReportingSenior') -and -not $_.legacyHidden})
NewScreen 'frmUXSeniors' 'Reporting seniors' 8000;$f=$script:swiftForm;UXChrome $f
$f.Controls('lblHeading').Top=[int](1200)
[void](Control $f 100 'lblInfo' 400 1800 12600 600 'Reusable reporting senior block details and manual RSCA baselines by evaluated paygrade. Saved reports change only when details are deliberately applied.')
[void](Control $f 110 'lstSeniors' 400 2700 12600 3200);$c=$script:swiftControl;$c.ColumnCount=3;$c.ColumnWidths='0;11500;0';$c.OnDblClick='=RSOpenSenior()'
[void](Control $f 104 'cmdNew' 400 6300 6000 550 'New reporting senior' '=RSOpenSenior(True)')
[void](Control $f 104 'cmdOpen' 7000 6300 6000 550 'Open selected record' '=RSOpenSenior()')
[void](Control $f 100 'lblStatus' 400 7150 12600 500 '')
$f.OnLoad='=RSLoadDirectory()';SaveScreen $f 'frmUXSeniors'
NewScreen 'frmUXSeniorEditor' 'Reporting senior record' 8400;$f=$script:swiftForm;UXChrome $f
$f.Controls('lblHeading').Top=[int](1100)
foreach($name in @('txtSeniorID','txtVersion')){[void](Control $f 109 $name 10 10 100 100);$script:swiftControl.Visible=$false}
[void](Control $f 100 'lblTitle' 400 1650 12600 300 'Record name')
[void](Control $f 109 'txtTitle' 400 2000 12600 400);$script:swiftControl.AfterUpdate='=RSPendingLabel()'
$tabs=@('Reporting senior details','Address & contact','Additional details','RSCA history')
for($page=0;$page -lt $tabs.Count;$page++){[void](Control $f 104 ('cmdPage'+$page) (400+$page*3200) 2600 3000 450 $tabs[$page] ('=RSSelectSeniorPage('+$page+')'))}
for($i=0;$i -lt $rsKeys.Count;$i++){
    $field=$rsKeys[$i];$page=[math]::Floor($i/6);$slot=$i%6;$x=400+($slot%2)*6500;$y=3300+[math]::Floor($slot/2)*1050
    [void](Control $f 100 ('lblRS'+$field.key) $x $y 6000 450 $field.label);$script:swiftControl.Tag='SR:'+$page
    [void](Control $f 109 ('rs'+$field.key) $x ($y+500) 6000 400);$c=$script:swiftControl;$c.Tag='SR:'+$page;$c.AfterUpdate='=RSPendingLabel()'
}
$inputs=@(@('cboPaygrade','Evaluated paygrade',111),@('txtAsOf','As-of date (YYYY-MM-DD)',109),@('txtRSCA','Historical RSCA',109),@('txtTraitCount','Previous graded trait count (not reports)',109),@('txtTraitSum','Exact trait sum (optional)',109),@('txtSource','Source / note',109))
for($i=0;$i -lt $inputs.Count;$i++){
    $item=$inputs[$i];$x=400+($i%2)*6500;$y=3250+[math]::Floor($i/2)*800
    [void](Control $f 100 ('lbl'+$item[0]) $x $y 6000 300 $item[1]);$script:swiftControl.Tag='SR:3'
    [void](Control $f $item[2] $item[0] $x ($y+350) 6000 350);$c=$script:swiftControl;$c.Tag='SR:3'
    if($item[0] -eq 'cboPaygrade'){$c.RowSourceType='Value List';$c.RowSource='E1;E2;E3;E4;E5;E6;E7;E8;E9;W1;W2;W3;W4;W5;O1;O2;O3;O4;O5;O6';$c.LimitToList=$true}
}
[void](Control $f 104 'cmdAddBaseline' 400 5800 12600 450 'Add historical baseline (save record details first)' '=RSAddBaselineUI()');$script:swiftControl.Tag='SR:3'
[void](Control $f 110 'lstBaselines' 400 6350 12600 600);$c=$script:swiftControl;$c.Tag='SR:3';$c.ColumnCount=6;$c.ColumnWidths='1200;1900;1700;1800;1800;4000';$c.ColumnHeads=$true
[void](Control $f 104 'cmdBack' 400 7300 6000 500 'Back to reporting seniors' '=UXNavigate("frmUXSeniors")')
[void](Control $f 104 'cmdSave' 7000 7300 6000 500 'Save record details' '=RSSaveSeniorUI()')
[void](Control $f 100 'lblStatus' 400 7900 12600 450 'Changes to this record do not automatically change saved reports. Exact trait sum avoids rounded-RSCA estimates.')
$f.OnLoad='=RSLoadSeniorEditor()';SaveScreen $f 'frmUXSeniorEditor'
NewScreen 'frmUXGroupRSCA' 'Group PSG RSCA' 8500;$f=$script:swiftForm;UXChrome $f
$f.Controls('lblHeading').Top=[int](1200)
[void](Control $f 109 'txtGroupID' 10 10 100 100);$script:swiftControl.Visible=$false
[void](Control $f 100 'lblSenior' 400 1800 12600 300 'Reporting senior record')
[void](Control $f 111 'cboSenior' 400 2200 12600 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;12000';$c.LimitToList=$true;$c.AfterUpdate='=RSGroupSeniorChanged()'
[void](Control $f 100 'lblBaseline' 400 2950 12600 300 'Historical RSCA baseline for this group paygrade')
[void](Control $f 111 'cboBaseline' 400 3350 12600 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;12000';$c.LimitToList=$true
[void](Control $f 106 'chkExcludes' 400 4250 350 350)
[void](Control $f 100 'lblExcludes' 850 4250 12000 650 'Historical totals exclude every report in this summary group. Confirm this to prevent double counting.')
[void](Control $f 104 'cmdSave' 400 5200 6000 550 'Save selection and calculate PSG RSCA' '=RSSaveGroupCalculation()')
[void](Control $f 104 'cmdStage' 7000 5200 6000 550 'Stage senior details for group members' '=RSStageGroupSenior()')
[void](Control $f 100 'lblCalculation' 400 6150 12600 500 'PSG RSCA: unavailable');$script:swiftControl.FontSize=14
[void](Control $f 100 'lblStatus' 400 6850 12600 700 '')
[void](Control $f 104 'cmdDirectory' 400 7800 6000 500 'Manage reporting seniors' '=RSOpenDirectory()')
[void](Control $f 104 'cmdBack' 7000 7800 6000 500 'Back to summary group' '=UXCloseScreen("frmUXGroupRSCA")')
$f.OnLoad='=RSLoadGroupCalculator()';SaveScreen $f 'frmUXGroupRSCA'
$app.DoCmd.OpenForm('frmUXResources',1);$f=$app.Forms.Item('frmUXResources')
[void](Control $f 104 'cmdSeniors' 400 6800 12600 550 'Reporting seniors and RSCA history' '=RSOpenDirectory()');SaveScreen $f 'frmUXResources'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
[void](Control $f 111 'uxSeniorRecord' 3100 6350 5800 400);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;5500';$c.LimitToList=$true;$c.RowSource='SELECT SeniorID,Title FROM ReportingSeniorRecord ORDER BY Title';$c.Tag='UX:1:0:Action'
[void](Control $f 104 'uxUseSenior' 9300 6350 4000 400 'Use record in pending draft' '=RSUseSeniorInEditor()');$script:swiftControl.Tag='UX:1:0:Action';SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
[void](Control $f 104 'uxRSCA' 400 7700 6000 500 'Reporting senior and PSG RSCA' '=RSOpenGroupCalculator()');$script:swiftControl.Tag='WBX:Members';SaveScreen $f 'frmMajorWorkbench'
