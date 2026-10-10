# Called after final workspace chrome and before compiling.
foreach($module in @('modUXBlockTools','modUXTraitGrid')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
NewScreen 'frmUXSharedFields' 'Shared blocks' 5000;$f=$script:swiftForm
$app.DeleteControl($f.Name,'lblHeading');$f.ScrollBars=2;$f.Width=10450
[void](Control $f 104 'anchor' 0 0 100 100 '');$script:swiftControl.TabStop=$false
foreach($field in $uxFields){
    if($field.key -notin $uxShared){continue}
    [void](Control $f 100 ('label'+$field.key) 150 150 3000 300 $field.label)
    $type=if($field.checkbox){106}elseif($field.options){111}else{109}
    [void](Control $f $type ('value'+$field.key) 150 500 3000 360);$c=$script:swiftControl
    if($field.checkbox){$c.TripleState=$true}elseif($field.options){$c.RowSourceType='Value List';$c.RowSource=$field.options -join ';';$c.LimitToList=$false}
    $c.AfterUpdate='=UXBlockEdit("'+$field.key+'")';if($type -ne 106){$c.OnChange='=UXInvalidateBulkPreview()'}
    $c.ControlTipText=$field.label
    [void](Control $f 104 ('clear'+$field.key) 3300 500 450 330 'X' ('=UXBlockClear("'+$field.key+'")'));$script:swiftControl.ControlTipText='Clear '+$field.label
    [void](Control $f 100 ('state'+$field.key) 150 900 3600 260 'Unchanged');$script:swiftControl.FontSize=8
}
SaveScreen $f 'frmUXSharedFields'
$app.DoCmd.OpenForm('frmUXBulk',1);$f=$app.Forms.Item('frmUXBulk')
foreach($c in $f.Controls){if($c.Tag -like 'BLX:*' -or $c.Name -match '^uxSection\dPage'){$c.Visible=$false}}
$f.Controls('lblHeading').Top=850;$f.Controls('uxGroupTitle').Top=1450;$f.Controls('uxSection').Top=2550
[void](Control $f 109 'uxKind' 0 0 100 100);$script:swiftControl.Visible=$false
[void](Control $f 112 'uxBlocks' 2850 3050 10450 3700);$script:swiftControl.SourceObject='Form.frmUXSharedFields';$script:swiftControl.BorderStyle=0
$i=0;foreach($section in @(0,1,2,6)){
    [void](Control $f 104 ('uxBlockSection'+$section) (2850+$i*2620) 2050 2500 400 $uxSections[$section] ('=UXBulkSection('+$section+')'));$script:swiftControl.FontSize=9;$i++
}
$f.Controls('uxStatus').Left=2850;$f.Controls('uxStatus').Top=6800;$f.Controls('uxStatus').Width=10450
$f.Controls('lblReason').Left=2850;$f.Controls('lblReason').Top=7450;$f.Controls('lblReason').Width=2400
$f.Controls('txtReason').Left=5350;$f.Controls('txtReason').Top=7400;$f.Controls('txtReason').Width=7950
$f.Controls('cmdPreview').Left=6350;$f.Controls('cmdPreview').Top=8050;$f.Controls('cmdPreview').Width=3300
$f.Controls('cmdApply').Left=9850;$f.Controls('cmdApply').Top=8050;$f.Controls('cmdApply').Width=3450
$f.Controls('cmdBack').Left=2850;$f.Controls('cmdBack').Top=8050;$f.Controls('cmdBack').Width=3300;$f.Controls('cmdBack').OnClick='=UXCloseBlocks()'
$f.Controls('lstPreview').Left=2850;$f.Controls('lstPreview').Top=3050;$f.Controls('lstPreview').Width=10450;$f.Controls('lstPreview').Height=1600;$f.Controls('lstPreview').Visible=$false
$f.Controls('txtFullValues').Left=2850;$f.Controls('txtFullValues').Top=4850;$f.Controls('txtFullValues').Width=10450;$f.Controls('txtFullValues').Height=1850;$f.Controls('txtFullValues').Visible=$false
$f.Section(0).Height=9000;$f.OnResize='=UXBlockViewport()'
SaveScreen $f 'frmUXBulk'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
[void](Control $f 104 'uxDefaultsEditor' 2850 3000 10450 550 'Edit group defaults in report block order' '=UXOpenGroupDefaults()');$script:swiftControl.Tag='WBX:Defaults';$script:swiftControl.Visible=$false
SaveScreen $f 'frmMajorWorkbench'
# Editable rows live in a per-session temp database, never the Report table.
NewScreen 'frmUXTraitRows' 'Trait comparison' 700;$f=$script:swiftForm
$app.DeleteControl($f.Name,'lblHeading');$f.DefaultView=1;$f.AllowAdditions=$false;$f.AllowDeletions=$false;$f.RecordSelectors=$false;$f.ScrollBars=2;$f.Width=10450
$app.DoCmd.RunCommand(37) # Form header/footer
$f.Section(1).Height=950;$f.Section(2).Height=0;$f.Section(0).Height=650
# CreateControl section argument is explicit for column headings.
function GridHeading($name,$x,$w,$caption){$c=$app.CreateControl($f.Name,100,1,'','',$x,20,$w,900);$c.Name=$name;$c.Caption=$caption;$c.FontName='Segoe UI';$c.FontSize=8}
GridHeading 'headingMember' 50 2000 'Member / state'
[void](Control $f 109 'Member' 50 30 2000 300);$script:swiftControl.ControlSource='Member';$script:swiftControl.Locked=$true;$script:swiftControl.TabStop=$false
[void](Control $f 109 'RowStatus' 50 340 2000 280);$script:swiftControl.ControlSource='RowStatus';$script:swiftControl.Locked=$true;$script:swiftControl.TabStop=$false;$script:swiftControl.FontSize=8
foreach($name in @('ReportID','RevisionID','ObservationStatus')){[void](Control $f 109 $name 0 0 100 100);$script:swiftControl.ControlSource=$name;$script:swiftControl.Visible=$false}
for($block=33;$block -le 39;$block++){
    $x=2200+($block-33)*820;GridHeading ('heading'+$block) $x 760 ('Trait '+$block)
    [void](Control $f 111 ('Trait'+$block) $x 40 760 360);$c=$script:swiftControl;$c.ControlSource='Trait'+$block;$c.RowSourceType='Value List';$c.RowSource='NOB;1.0;2.0;3.0;4.0;5.0';$c.LimitToList=$false;$c.AfterUpdate='=UXGridQueue("Trait'+$block+'")'
}
GridHeading 'headingITA' 8100 700 'ITA'
[void](Control $f 109 'ITA' 8100 40 700 360);$script:swiftControl.ControlSource='ITA';$script:swiftControl.Locked=$true;$script:swiftControl.TabStop=$false
GridHeading 'headingPromotion' 9000 1350 'Promotion recommendation'
[void](Control $f 111 'PromotionRecommendation' 9000 40 1350 360);$c=$script:swiftControl;$c.ControlSource='PromotionRecommendation';$c.RowSourceType='Value List';$c.RowSource=';NOB;Significant Problems;Progressing;Promotable;Must Promote;Early Promote';$c.LimitToList=$true;$c.AfterUpdate='=UXGridQueue("PromotionRecommendation")'
foreach($c in $f.Controls){$c.Tag=[string]$c.Left+':'+[string]$c.Width}
$f.OnCurrent='=UXGridCurrent()'
SaveScreen $f 'frmUXTraitRows'
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison')
foreach($c in $f.Controls){if($c.Name -match '^(cboTrait|lblTrait)' -or $c.Name -in @('lstTraits','uxMember','uxITA','cboPromotion','cmdStage')){$c.Visible=$false}}
[void](Control $f 112 'uxTraitGrid' 2850 2200 10450 4600);$script:swiftControl.SourceObject='Form.frmUXTraitRows';$script:swiftControl.BorderStyle=0
[void](Control $f 100 'uxGridAverage' 2850 6950 10450 450 '')
[void](Control $f 109 'uxProposalDetail' 2850 5500 10450 1700);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2;$script:swiftControl.Visible=$false
$f.Controls('lstProposals').ColumnCount=6;$f.Controls('lstProposals').ColumnWidths='0;2600;2800;1800;1800;1100';$f.Controls('lstProposals').Height=2400;$f.Controls('lstProposals').OnClick='=UXGridDetail()'
$f.Controls('cmdBack').OnClick='=UXCloseComparison()'
$f.OnTimer='=UXGridDrain()';$f.TimerInterval=0
$f.OnClose='=UXGridClose()';$f.OnResize='=UXGridViewport()'
# build.ps1 attaches the actual unload class procedure after the generated
# forms are persisted and reopened; avoid an intermediate dangling event.
SaveScreen $f 'frmUXComparison'
