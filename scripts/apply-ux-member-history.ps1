if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXDirectory',(Join-Path $root 'src/modUXDirectory.bas'))}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modMemberHistory',(Join-Path $root 'src/modMemberHistory.bas'))}
$app.DoCmd.OpenForm('frmUXDirectory',1);$f=$app.Forms.Item('frmUXDirectory')
[void](Control $f 104 'cmdHistory' 400 6600 12600 500 'Written reports and prior evaluations for selected member' '=UXDirectoryHistory()');$f.Section(0).Height=7600;SaveScreen $f 'frmUXDirectory'
NewScreen 'frmUXMemberHistory' 'Member evaluation history' 7600;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblMember' 400 950 12600 400 '')
[void](Control $f 100 'lblScope' 400 1450 12600 500 'Recorded prior evaluations and saved Swift reports. Group prior columns use the latest ending before the current report starts; equal ending dates require review.')
[void](Control $f 110 'lstHistory' 400 2150 12600 3300);$c=$script:swiftControl;$c.ColumnCount=6;$c.ColumnWidths='0;2300;1800;4400;1500;2400'
[void](Control $f 100 'lblStatus' 400 5600 12600 550 '')
[void](Control $f 104 'cmdAdd' 400 6400 4000 500 'Record prior evaluation' '=UXPriorOpen(True)')
[void](Control $f 104 'cmdEdit' 4650 6400 4100 500 'Edit selected prior record' '=UXPriorOpen(False)')
[void](Control $f 104 'cmdClose' 9000 6400 4000 500 'Back to members' '=UXDirectoryClose("frmUXMemberHistory")')
SaveScreen $f 'frmUXMemberHistory'
NewScreen 'frmUXPriorEditor' 'Record a prior evaluation' 10200;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.ScrollBars=2
$fields=@(@('cboKind','Report type','EVAL;CHIEFEVAL;FITREP'),@('txtRank','Rank/rate',''),@('txtFrom','Report from date',''),@('txtTo','Report to date',''),@('cboPromotion','Promotion recommendation','Significant Problems;Progressing;Promotable;Must Promote;Early Promote'),@('cboObservation','Observation','Observed;NOB'))
$y=950;foreach($item in $fields){[void](Control $f 100 ('lbl'+$item[0]) 400 $y 4100 350 $item[1]);$type=if($item[2]){111}else{109};[void](Control $f $type $item[0] 4600 $y 8000 450);$c=$script:swiftControl;if($type -eq 111){$c.RowSourceType='Value List';$c.RowSource=$item[2];$c.LimitToList=$true};$y+=600}
$f.Controls('cboKind').AfterUpdate='=UXPriorLabels()'
foreach($b in 33..39){[void](Control $f 100 ('lblTrait'+$b) 400 $y 8300 350 ('Trait '+$b));[void](Control $f 111 ('cboTrait'+$b) 9000 $y 3600 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='1.0;2.0;3.0;4.0;5.0;NOB';$c.LimitToList=$true;$y+=550}
[void](Control $f 100 'lblSource' 400 $y 12600 350 'Source/context (required; 255 characters)');$y+=400
[void](Control $f 109 'txtSource' 400 $y 12600 450);$y+=650
[void](Control $f 100 'lblStatus' 400 $y 12600 500 'Unentered traits remain incomplete. This record does not change the current report.');$y+=650
[void](Control $f 104 'cmdSave' 400 $y 6000 500 'Save prior evaluation' '=UXPriorSave()')
[void](Control $f 104 'cmdCancel' 6700 $y 6300 500 'Cancel' '=UXDirectoryClose("frmUXPriorEditor")')
$f.Section(0).Height=[int]($y+1100);SaveScreen $f 'frmUXPriorEditor'
