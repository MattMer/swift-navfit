# This startup notice has no VBA events and can open while content is disabled.
NewScreen 'frmEnableContent' 'Enable Swift NAVFIT content' 4500;$f=$script:swiftForm
[void](Control $f 100 'lblInstructions' 500 1400 11500 1600 "Access has disabled this app's code. Click Enable Content in the yellow security bar, then close and reopen this file. To enable future preview builds, add only the output\ux preview folder under File > Options > Trust Center > Trust Center Settings > Trusted Locations.")
[void](Control $f 100 'lblExplanation' 500 3300 11500 700 'Your reports are unchanged. Application features require enabled content.')
SaveScreen $f 'frmEnableContent'
# Native section surfaces and focused destinations. Called after shared chrome.
$surfaceModule=[Collections.Generic.List[string]]::new()
$surfaceModule.Add('Attribute VB_Name = "modUXSurfaceFields"');$surfaceModule.Add('Option Explicit')
function SurfaceProperty($functionName,$type,$property,$default){
    $surfaceModule.Add('Public Function '+$functionName+'(ByVal key As String) As '+$type)
    $surfaceModule.Add('Select Case key')
    for($i=0;$i -lt $uxFields.Count;$i++){
        $field=$uxFields[$i]
        $value=if($property -eq 'control'){'"'+$uxNames[$i]+'"'}elseif($type -eq 'String'){'"'+([string]$field.$property).Replace('"','""')+'"'}elseif($type -eq 'Boolean'){if($field.$property){'True'}else{'False'}}else{if($null -eq $field.$property){'0'}else{[string]$field.$property}}
        if($property -eq 'options'){$value='"'+(($field.options -join ';').Replace('"','""'))+'"'}
        $surfaceModule.Add('Case "'+$field.key+'": '+$functionName+' = '+$value)
    }
    $surfaceModule.Add('Case Else: '+$functionName+' = '+$default);$surfaceModule.Add('End Select');$surfaceModule.Add('End Function')
}
SurfaceProperty 'UXFieldControl' 'String' 'control' '""'
SurfaceProperty 'UXFieldSection' 'Long' 'section' '7'
SurfaceProperty 'UXFieldLimit' 'Long' 'limit' '0'
SurfaceProperty 'UXFieldDigits' 'Long' 'digits' '0'
SurfaceProperty 'UXFieldDate' 'Boolean' 'date' 'False'
SurfaceProperty 'UXFieldFlag' 'Boolean' 'checkbox' 'False'
SurfaceProperty 'UXFieldMultiline' 'Boolean' 'multiline' 'False'
SurfaceProperty 'UXFieldLegacy' 'Boolean' 'legacyHidden' 'False'
SurfaceProperty 'UXFieldOptions' 'String' 'options' '""'
$surfaceModule.Add('Public Function UXFieldApplies(ByVal key As String, ByVal kind As String) As Boolean')
$surfaceModule.Add('UXFieldApplies = True: Select Case key')
foreach($field in $uxFields|Where-Object forms){$surfaceModule.Add('Case "'+$field.key+'": UXFieldApplies = ('+(($field.forms|ForEach-Object{'kind = "'+$_+'"'})-join ' Or ')+')')}
$surfaceModule.Add('End Select: End Function')
$surfaceModule.Add('Public Function UXFieldCaption(ByVal key As String, ByVal kind As String) As String')
$surfaceModule.Add('Dim label As String, block As String: label = UXFieldLabel(key): Select Case key')
foreach($field in $uxFields){
    $surfaceModule.Add('Case "'+$field.key+'": block = "'+$field.block+'"')
    if($field.blocks){foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){$surfaceModule.Add('If kind = "'+$kind+'" Then block = "'+$field.blocks.$kind+'"')}}
}
$surfaceModule.Add('End Select')
$surfaceModule.Add('If key = "Designator" Then block = "3"')
$surfaceModule.Add('If key = "RankRate" And kind = "CHIEFEVAL" Then label = "Rating"')
$surfaceModule.Add('If IsTraitKey(key) Then label = Replace(Replace(CatalogueTraitTitle(kind, CLng(Mid$(key, 6))), vbCr, ""), vbLf, " ")')
$surfaceModule.Add('If Len(block) > 0 Then label = block & ". " & label')
$surfaceModule.Add('UXFieldCaption = label: End Function')
$surfaceModule.Add('Public Function UXSectionKeys(ByVal section As Long, ByVal kind As String) As Variant')
$surfaceModule.Add('Select Case kind & ":" & section')
foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){for($section=0;$section -le 7;$section++){
    $ordered=@($uxFields|Where-Object {$_.section -eq $section}|Sort-Object @{Expression={ $block=if($_.blocks){$_.blocks.$kind}else{$_.block};if($_.key -eq 'Designator'){$block='3'};if($block -match '^\d+'){[int]$Matches[0]}else{999}}})
    $surfaceModule.Add('Case "'+$kind+':'+$section+'": UXSectionKeys = Array('+(($ordered|ForEach-Object{'"'+$_.key+'"'})-join ',')+')')
}}
$surfaceModule.Add('Case Else: UXSectionKeys = Array(): End Select: End Function')
$surfaceModulePath=Join-Path $root 'build/modUXSurfaceFields.bas'
[IO.File]::WriteAllText($surfaceModulePath,($surfaceModule -join "`r`n"),[Text.UTF8Encoding]::new($false))
$app.LoadFromText(5,'modUXSurfaceFields',$surfaceModulePath)
$app.LoadFromText(5,'modUXSurface',(Join-Path $root 'src/modUXSurface.bas'))
$app.LoadFromText(5,'modUXDashboard',(Join-Path $root 'src/modUXDashboard.bas'))
NewScreen 'frmUXEditorFields' 'Report section' 12500;$f=$script:swiftForm
$f.Controls('lblHeading').Visible=$false;$f.Width=[int](8000);$f.ScrollBars=2;$f.BorderStyle=0;$f.Section(0).BackColor=16777215
[void](Control $f 109 'uxFocusAnchor' 0 0 20 20);$c=$script:swiftControl;$c.Locked=$true;$c.TabStop=$false;$c.BorderStyle=0
for($i=0;$i -lt $uxFields.Count;$i++){
    $field=$uxFields[$i];if($field.key -eq 'FormKind'){continue}
    $type=if($field.checkbox){106}elseif($field.options){111}else{109}
    [void](Control $f $type $uxNames[$i] 200 200 3600 450);$c=$script:swiftControl;$c.Tag='SURF:'+$field.key;$c.AfterUpdate='=UXSurfaceEdited("'+$field.key+'")'
    if($field.checkbox){$c.TripleState=$true}elseif($field.options){$c.RowSourceType='Value List';$c.RowSource=$field.options -join ';';$c.LimitToList=$false}
    if($field.multiline){$c.EnterKeyBehavior=$true;$c.ScrollBars=2}
    if($field.key -match '^Trait(\d+)$'){$c.OnEnter='=UXSurfaceTrait('+$Matches[1]+')'}
    $label=$app.CreateControl($f.Name,100,0,$c.Name,'',200,100,3600,450);$label.Name='label'+$field.key;$label.Caption=$field.label;$label.FontName='Segoe UI';$label.FontSize=10
    [void](Control $f 100 ('err'+$field.key) 200 900 3600 250 '');$script:swiftControl.ForeColor=2893478;$script:swiftControl.FontSize=9
}
[void](Control $f 109 'uxTraitGuidance' 200 9500 7500 1800);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2;$c.BackColor=16381429
[void](Control $f 100 'uxNarrativeCount' 200 11000 7500 350 '')
[void](Control $f 104 'cmdSnippets' 200 11500 3000 450 'Insert narrative snippet' '=OpenScreen("frmSnippets")')
SaveScreen $f 'frmUXEditorFields'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
foreach($c in $f.Controls){if($c.Tag -like 'UX:*'){$c.Visible=$false;$c.Tag='BUFFER'} }
foreach($name in $uxNames){$f.Controls($name).Visible=$false}
$f.Controls('cboKind').Visible=$true;$f.Controls('cboKind').AfterUpdate='=UXEditorKindChanged()'
$f.Section(0).Height=[int](8650)
$f.Controls('uxReport').Left=[int](2850);$f.Controls('uxReport').Top=[int](850);$f.Controls('uxReport').Width=[int](6200);$f.Controls('uxReport').Height=[int](650);$f.Controls('uxReport').FontSize=13
$f.Controls('uxITA').Left=[int](2850);$f.Controls('uxITA').Top=[int](1530);$f.Controls('uxITA').Width=[int](10500);$f.Controls('uxITA').Height=[int](400)
$f.Controls('uxSGA').Left=[int](2850);$f.Controls('uxSGA').Top=[int](1950);$f.Controls('uxSGA').Width=[int](8500);$f.Controls('uxSGA').Height=[int](400)
$f.Controls('uxGroup').Left=[int](11350);$f.Controls('uxGroup').Top=[int](1950);$f.Controls('uxGroup').Width=[int](2000)
$f.Controls('uxSection').Left=[int](5250);$f.Controls('uxSection').Top=[int](2450);$f.Controls('uxSection').Width=[int](8050);$f.Controls('uxSection').Height=[int](450)
$f.Controls('uxError').Left=[int](5250);$f.Controls('uxError').Top=[int](2930);$f.Controls('uxError').Width=[int](8050);$f.Controls('uxError').Height=[int](300)
$f.Controls('uxLocked').Left=[int](2850);$f.Controls('uxLocked').Top=[int](2350);$f.Controls('uxLocked').Width=[int](2350);$f.Controls('uxLocked').Height=[int](600)
for($i=0;$i -lt $uxSections.Count;$i++){$c=$f.Controls('uxSection'+$i);$c.Left=[int](2800);$c.Top=[int](3050+$i*430);$c.Width=[int](2150);$c.Height=[int](400);$c.FontSize=9;$c.Caption=$uxSections[$i].Replace('&','&&')}
[void](Control $f 112 'uxFieldSurface' 5200 3650 8100 3700);$script:swiftControl.SourceObject='Form.frmUXEditorFields';$script:swiftControl.BorderStyle=0
# Access: 0 = left/top, 1 = right/bottom, 2 = stretch. Keep this
# fixed laptop layout stable until the responsive layout pass.
$script:swiftControl.HorizontalAnchor=0;$script:swiftControl.VerticalAnchor=0
[void](Control $f 110 'uxErrorList' 5250 3200 8050 430);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;7900';$c.OnClick='=UXSurfaceErrorSelected()';$c.Visible=$false
foreach($name in @('uxPrevious','uxPage','uxNext')){$f.Controls($name).Visible=$false}
foreach($name in @('cmdSave','cmdClose','uxDirty','lblSaved','lblRecovery')){$f.Controls($name).Top=[int](7750)}
foreach($name in @('cmdSave','cmdClose','uxDirty','lblSaved','lblRecovery')){$f.Controls($name).VerticalAnchor=0}
foreach($name in @('cmdSave','cmdClose')){$f.Controls($name).Height=[int](500)}
$f.Controls('cmdClose').Left=[int](5250);$f.Controls('cmdClose').Width=[int](2350)
$f.Controls('cmdSave').Left=[int](10800);$f.Controls('cmdSave').Width=[int](2500);$f.Controls('cmdSave').Caption='Save draft';$f.Controls('cmdSave').BackColor=11230215;$f.Controls('cmdSave').ForeColor=16777215
$f.Controls('uxDirty').Left=[int](7800);$f.Controls('uxDirty').Width=[int](2800)
$f.Controls('lblSaved').Top=[int](7350);$f.Controls('lblSaved').Left=[int](5250);$f.Controls('lblSaved').Width=[int](8050)
$f.Controls('lblRecovery').Top=[int](8350);$f.Controls('lblRecovery').Left=[int](5250);$f.Controls('lblRecovery').Width=[int](8050)
foreach($name in @('cmdRoute','cmdComments','cmdChecklist','cmdCompare','cmdRestore','lstRevisions')){$f.Controls($name).Tag='LOCAL:'+$(if($name -in @('cmdCompare','cmdRestore','lstRevisions')){'9'}else{'8'})}
$f.Controls('lstRevisions').Left=[int](5250);$f.Controls('lstRevisions').Top=[int](4800);$f.Controls('lstRevisions').Width=[int](8050);$f.Controls('lstRevisions').Height=[int](2400)
foreach($entry in @(@('cmdRoute',3700),@('cmdComments',4300),@('cmdChecklist',4900),@('cmdCompare',3700),@('cmdRestore',4300))){$c=$f.Controls($entry[0]);$c.Left=[int](5250);$c.Top=[int]($entry[1]);$c.Width=[int](3900);$c.Height=[int](500)}
$f.Controls('uxSeniorRecord').Left=[int](5250);$f.Controls('uxSeniorRecord').Top=[int](3150);$f.Controls('uxSeniorRecord').Width=[int](4800);$f.Controls('uxSeniorRecord').Tag='LOCAL:0'
$f.Controls('uxUseSenior').Left=[int](10250);$f.Controls('uxUseSenior').Top=[int](3150);$f.Controls('uxUseSenior').Width=[int](3050);$f.Controls('uxUseSenior').Tag='LOCAL:0'
$f.OnResize='=UXResizeEditor()'
# Class unload handler is bound in build.ps1 after all controls are persisted.
SaveScreen $f 'frmEditor'
# Overview: real metrics and recent reports, not a launch-button grid.
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome')
foreach($c in $f.Controls){if($c.Name -match '^uxCard' -or $c.Name -in @('cmdRecovery','cmdImportReview','cmdPDF')){$c.Visible=$false}}
$f.Controls('lblHeading').Top=[int](1000);$f.Controls('lblHeading').Height=[int](500)
$cards=@(@('cmdLibrary','Reports','uxReportCount'),@('cmdMajor','Summary groups','uxGroupCount'),@('cmdMyWork','My Work','uxWorkCount'))
for($i=0;$i -lt 3;$i++){
    $x=2850+$i*3500;$c=$f.Controls($cards[$i][0]);$c.Left=[int]($x);$c.Top=[int](1850);$c.Width=[int](3200);$c.Height=[int](550)
    [void](Control $f 100 $cards[$i][2] $x 2500 3200 550 '0');$script:swiftControl.FontSize=26
}
[void](Control $f 100 'uxRecoveryStatus' 2850 3250 7300 500 'No recovery checkpoints.')
$f.Controls('cmdRecovery').Left=[int](10500);$f.Controls('cmdRecovery').Top=[int](3200);$f.Controls('cmdRecovery').Width=[int](2800);$f.Controls('cmdRecovery').Height=[int](500);$f.Controls('cmdRecovery').Visible=$true
[void](Control $f 100 'uxRecentCaption' 2850 4000 10450 400 'Recent reports');$script:swiftControl.FontSize=16
[void](Control $f 110 'uxRecentReports' 2850 4500 10450 2300);$c=$script:swiftControl;$c.ColumnCount=6;$c.ColumnHeads=$true;$c.ColumnWidths='0;3000;1100;2200;2200;1500';$c.OnDblClick='=UXOpenRecentReport()'
[void](Control $f 104 'uxOpenRecent' 2850 7100 3000 500 'Open selected report' '=UXOpenRecentReport()')
[void](Control $f 104 'uxNewReport' 9850 7100 3450 500 'New report' '=UXNewReport()')
$f.Section(0).Height=[int](8200)
SaveScreen $f 'frmHome'
# Imports are contextual Reports actions, rather than dashboard destinations.
$app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary')
[void](Control $f 104 'uxImport' 10500 900 2800 500 'Import NAVFIT database' '=ChooseReviewFile()')
[void](Control $f 104 'uxPDF' 2850 8100 3000 450 'Import structured PDF' '=OpenScreen("frmPDFImport")')
$f.Controls('lblHeading').Top=[int](800);$f.Controls('lblHeading').Width=[int](7000);$f.Controls('lblHeading').Height=[int](400)
foreach($entry in @(@('uxTypeFilter','Report type','All;EVAL;CHIEFEVAL;FITREP',2850),@('uxStatusFilter','Local review status','All;Draft;In review;Review needs restart;Local approval;Returned',6100))){
    [void](Control $f 100 ('lbl'+$entry[0]) $entry[3] 2200 2900 300 $entry[1])
    [void](Control $f 111 $entry[0] $entry[3] 2530 2900 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource=$entry[2];$c.DefaultValue='"All"';$c.AfterUpdate='=RefreshLibrary()';$c.LimitToList=$true
}
[void](Control $f 104 'uxClearFilters' 9850 2530 3450 400 'Clear filters' '=UXClearReportFilters()')
$f.Controls('lstReports').Top=[int](3250);$f.Controls('lstReports').Height=[int](2200)
$f.Controls('cmdNew').OnClick='=UXNewReport()'
SaveScreen $f 'frmLibrary'
NewScreen 'frmUXNewReport' 'New report' 3500;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.Width=[int](10500)
$f.Controls('lblHeading').Width=[int](9600)
[void](Control $f 100 'lblInfo' 400 1000 9600 650 'Choose a report type. Its official block captions and trait guidance will be used in the editor.')
foreach($entry in @(@('EVAL','E1-E6 EVAL',400),@('CHIEFEVAL','Chief EVAL',3750),@('FITREP','FITREP',7100))){[void](Control $f 104 ('cmd'+$entry[0]) $entry[2] 2000 3000 550 $entry[1] ('=UXChooseNewType("'+$entry[0]+'")'))}
[void](Control $f 104 'cmdCancel' 400 2800 9600 450 'Cancel' '=UXChooseNewType("")')
SaveScreen $f 'frmUXNewReport'
NewScreen 'frmUXNewGroup' 'New summary group' 5200;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true
[void](Control $f 100 'lblTitle' 400 1100 4000 350 'Summary group name')
[void](Control $f 109 'txtTitle' 4500 1050 8200 450)
[void](Control $f 100 'lblKind' 400 1950 4000 350 'Report type')
[void](Control $f 111 'cboKind' 4500 1900 8200 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='EVAL;CHIEFEVAL;FITREP';$c.DefaultValue='"EVAL"';$c.LimitToList=$true
[void](Control $f 100 'lblPaygrade' 400 2800 4000 350 'Member paygrade')
[void](Control $f 111 'cboPaygrade' 4500 2750 8200 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='E1;E2;E3;E4;E5;E6;E7;E8;E9;W1;W2;W3;W4;W5;O1;O2;O3;O4;O5;O6';$c.LimitToList=$true
[void](Control $f 104 'cmdCreate' 7000 3650 5700 550 'Create group' '=UXCreateGroupUI()')
[void](Control $f 104 'cmdCancel' 400 3650 5700 550 'Cancel' '=UXCloseScreen("frmUXNewGroup")')
[void](Control $f 100 'lblStatus' 400 4500 12300 500 '')
SaveScreen $f 'frmUXNewGroup'
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups');$f.Controls('cmdNew').OnClick='=UXNewGroup()';SaveScreen $f 'frmUXGroups'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
$f.Section(0).Height=[int](8600);$f.ScrollBars=0
[void](Control $f 100 'uxGroupTitle' 2850 900 10450 550 'Summary group');$script:swiftControl.FontSize=20
[void](Control $f 100 'uxGroupMeta' 2850 1500 10450 450 '')
for($i=0;$i -lt 5;$i++){$c=$f.Controls('uxMode'+$i);$c.Left=[int](2850+$i*2100);$c.Top=[int](2100);$c.Width=[int](1900);$c.Height=[int](500)}
$f.Controls('lstReports').Left=[int](2850);$f.Controls('lstReports').Top=[int](3000);$f.Controls('lstReports').Width=[int](10450);$f.Controls('lstReports').Height=[int](2900)
$f.Controls('uxMembers').Left=[int](2850);$f.Controls('uxMembers').Top=[int](6300);$f.Controls('uxMembers').Width=[int](3300)
$f.Controls('uxBulk').Left=[int](6350);$f.Controls('uxBulk').Top=[int](6300);$f.Controls('uxBulk').Width=[int](3300)
$f.Controls('uxRSCA').Left=[int](9850);$f.Controls('uxRSCA').Top=[int](6300);$f.Controls('uxRSCA').Width=[int](3450)
$f.Controls('uxComparison').Left=[int](2850);$f.Controls('uxComparison').Top=[int](3000);$f.Controls('uxComparison').Width=[int](10450)
$f.Controls('uxBack').Left=[int](2850);$f.Controls('uxBack').Top=[int](7600);$f.Controls('uxBack').Width=[int](3450)
$f.Controls('lblStatus').Left=[int](2850);$f.Controls('lblStatus').Top=[int](6900);$f.Controls('lblStatus').Width=[int](10450);$f.Controls('lblStatus').Height=[int](500)
$f.Controls('cboScope').DefaultValue='"Group"'
$f.Controls('cboField').ColumnCount=2;$f.Controls('cboField').ColumnWidths='0;3000'
[void](Control $f 104 'uxOpenGroupReport' 6350 7600 3300 500 'Open selected report' '=UXOpenGroupReport()');$script:swiftControl.Tag='WBX:Members:Review'
SaveScreen $f 'frmMajorWorkbench'
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison');$f.Section(0).Height=[int](8700);$f.ScrollBars=0
$f.Controls('lstTraits').Top=[int](2100);$f.Controls('lstTraits').Height=[int](2300)
$f.Controls('uxMember').Top=[int](4550);$f.Controls('uxMember').Height=[int](450)
for($block=33;$block -le 39;$block++){$f.Controls('lblTrait'+$block).Top=[int](5200);$f.Controls('cboTrait'+$block).Top=[int](5550)}
$f.Controls('uxITA').Top=[int](6100)
$f.Controls('cboPromotion').Top=[int](6600);$f.Controls('cmdStage').Top=[int](6600)
$f.Controls('lstProposals').Top=[int](2800);$f.Controls('lstProposals').Height=[int](3800)
$f.Controls('cmdApply').Top=[int](7500);$f.Controls('cmdBack').Top=[int](7500)
$f.Controls('lblStatus').Top=[int](8150);$f.Controls('lblStatus').Height=[int](500)
[void](Control $f 104 'uxMarksPage' 2850 1650 4800 400 'Trait marks' '=UXComparisonPage("Marks")')
[void](Control $f 104 'uxPreviewPage' 8100 1650 5200 400 'Preview staged changes' '=UXComparisonPage("Preview")')
SaveScreen $f 'frmUXComparison'
$app.DoCmd.OpenForm('frmMyWork',1);$f=$app.Forms.Item('frmMyWork')
foreach($c in $f.Controls){if($c.Name -notmatch '^uxNav' -and $c.Name -notin @('uxContext','uxBrand','uxRail','lblHeading')){$c.Tag='MWX:body'}}
$f.Section(0).Height=[int](8500);$f.ScrollBars=0
$f.Controls('lblHeading').Caption='My Work';$f.Controls('lblHeading').Top=[int](950)
$tabs=@(@('Assignments','Assigned work'),@('Assign','Assign review'),@('Contributors','Contributors'))
for($i=0;$i -lt 3;$i++){[void](Control $f 104 ('uxWorkTab'+$i) (2850+$i*3500) 1850 3200 500 $tabs[$i][1] ('=UXMyWorkMode("'+$tabs[$i][0]+'")'))}
$positions=@(@('chkAll',2850,2500,300,300),@('lblAll',3300,2450,7300,400),@('cmdRefresh',11000,2450,2300,450),@('lstTasks',2850,3000,10450,1700),@('txtHistory',2850,4850,10450,1300),@('lblNotes',2850,6300,10450,300),@('txtNotes',2850,6650,10450,550),@('cmdDraft',2850,7500,2400,500),@('cmdAccept',5500,7500,2400,500),@('cmdReturn',8150,7500,2400,500),@('cmdCancel',10800,7500,2500,500),@('lblAssignment',2850,3550,10450,550),@('lblReviewer',2850,4300,4800,350),@('cboReviewer',2850,4750,4800,450),@('lblDue',8100,4300,5200,350),@('txtDue',8100,4750,5200,450),@('cmdAssign',2850,7500,10450,500),@('cboContributor',2850,2900,7000,450),@('cmdSwitch',10100,2900,3200,450),@('txtContributor',2850,3850,7000,450),@('cmdAdd',10100,3850,3200,450),@('lblScope',2850,4750,10450,1600))
foreach($p in $positions){$c=$f.Controls($p[0]);$c.Left=[int]$p[1];$c.Top=[int]$p[2];$c.Width=[int]$p[3];$c.Height=[int]$p[4]}
$f.Controls('cmdDraft').Caption='Open selected report';$f.Controls('lstTasks').ColumnWidths='0;3300;1700;1300;1700;1500;700'
[void](Control $f 100 'uxAssignCaption' 2850 2500 10450 350 'Choose a saved report')
$script:swiftControl.Tag='MWX:body'
[void](Control $f 111 'uxAssignReport' 2850 2900 10450 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;10200';$c.AfterUpdate='=UXChooseAssignmentReport()';$c.Tag='MWX:body';$c.LimitToList=$true
$f.OnLoad='=UXLoadMyWork()'
SaveScreen $f 'frmMyWork'

# Shared compact navigation and viewport events, after final authored geometry.
$app.LoadFromText(5,'modUXViewport',(Join-Path $root 'src/modUXViewport.bas'))
foreach($screenName in $uxPages){
    $app.DoCmd.OpenForm($screenName,1);$f=$app.Forms.Item($screenName)
    [void](Control $f 111 'uxCompactNav' 250 150 2500 400);$c=$script:swiftControl
    $c.Visible=$false;$c.RowSourceType='Value List';$c.RowSource='frmHome;Overview;frmLibrary;Reports;frmUXGroups;Summary groups;frmMyWork;My Work;frmUXResources;Resources;frmUXSettings;Settings & Help'
    $c.ColumnCount=2;$c.ColumnWidths='0;2500';$c.LimitToList=$true;$c.TabIndex=0;$c.ControlTipText='Workspace navigation';$c.AfterUpdate='=UXCompactNavigate("'+$screenName+'")'
    if($screenName -eq 'frmEditor'){
        [void](Control $f 111 'uxCompactSection' 250 2650 5000 400);$c=$script:swiftControl;$c.Visible=$false;$c.RowSourceType='Value List'
        $c.RowSource=(0..9|Where-Object {$_ -ne 1}|ForEach-Object{[string]$_+';'+$uxSections[$_]}) -join ';';$c.ColumnCount=2;$c.ColumnWidths='0;5000';$c.LimitToList=$true;$c.TabIndex=7;$c.ControlTipText='Report section';$c.AfterUpdate='=UXChooseSectionCompact()'
        [void](Control $f 100 'uxAlias' 5900 2350 6000 350 '')
        # build.ps1 binds the unload class procedure after all form design edits
        # have been saved and the database reopened. Do not persist a dangling
        # event-procedure reference during intermediate construction.
    }else{$f.OnResize='=UXResizeWorkspace("'+$screenName+'")'}
    SaveScreen $f $screenName
}
