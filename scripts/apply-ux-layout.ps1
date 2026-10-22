if(-not $script:swiftManagedBuild){
 . (Join-Path $PSScriptRoot 'apply-ux-block28-schema.ps1')
 . (Join-Path $PSScriptRoot 'apply-ux-field-catalogue.ps1')
}
function UXChrome($f){
    $f.Width=[int](13800);$f.Section(0).BackColor=16316664;$f.OnActivate="=UXLoadContext()"
    $nav=@(@(([string]$app.Eval('UXWorkspaceRouteList()')).Split(';')) | ForEach-Object { ,($_.Split('|')) })
    for($i=0;$i -lt $nav.Count;$i++){[void](Control $f 104 ('uxNav'+$i) (300+$i*2210) 200 2100 450 $nav[$i][0] ('=UXNavigate("'+$nav[$i][1]+'")'))}
    [void](Control $f 100 'uxContext' 350 720 13000 300 'Swift NAVFIT | Local workspace')
}
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
foreach($c in $f.Controls){$c.Visible=$false;$c.Top=[int](10);$c.Left=[int](10);$c.Width=[int](100);$c.Height=[int](100)}
$f.Section(0).Height=[int](8200);$f.ScrollBars=0;UXChrome $f
[void](Control $f 100 'uxReport' 3100 1050 10300 350 'Report workspace')
[void](Control $f 100 'uxITA' 3100 1450 10300 350 'ITA - pending')
[void](Control $f 100 'uxSGA' 3100 1850 7600 500 'Summary Group Average - saved')
[void](Control $f 111 'uxGroup' 10800 1850 2500 400);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;2400';$c.AfterUpdate='=UXRefreshEditor()'
[void](Control $f 100 'uxSection' 3100 2450 10000 420 'Member & report details');$script:swiftControl.FontSize=16
[void](Control $f 100 'uxError' 3100 2900 10000 600 '');$script:swiftControl.ForeColor=255
[void](Control $f 100 'uxLocked' 3100 3550 10000 300 '')
for($i=0;$i -lt $uxSections.Count;$i++){[void](Control $f 104 ('uxSection'+$i) 350 (1200+$i*480) 2450 430 $uxSections[$i] ('=UXSelectSection('+$i+')'))}
$uxCounters=@{}
for($i=0;$i -lt $uxFields.Count;$i++){
    $field=$uxFields[$i];$name=$uxNames[$i]
    if($field.key -eq 'FormKind' -or $field.legacyHidden){continue}
    $section=[int]$field.section;if(-not $uxCounters.ContainsKey($section)){$uxCounters[$section]=0};$ordinal=$uxCounters[$section];$uxCounters[$section]++
    $page=[math]::Floor($ordinal/6);$slot=$ordinal%6;if($field.multiline){$page=1;$slot=0}elseif($section -eq 5){$page=0};$x=3100+($slot%2)*5100;$y=3900+[math]::Floor($slot/2)*800
    $tag='UX:'+$section+':'+$page+':'+$field.key
    $label=$field.label;if($field.block){$label=$field.block+'. '+$label}
    if($field.key -match '^Trait(\d+)$'){$trait=$uxCatalogue.Traits|Where-Object {$_.FormKind -eq 'EVAL' -and $_.Block -eq [int]$Matches[1]}|Select-Object -First 1;$label=$Matches[1]+'. '+($trait.Title -replace '\s+',' ')}
    [void](Control $f 100 ('uxLabel'+$field.key) $x $y 4900 320 $label);$script:swiftControl.Tag=$tag
    if($field.checkbox){
        [void](Control $f 106 ('uxChk'+$field.key) $x ($y+350) 320 320);$c=$script:swiftControl;$c.TripleState=$true;$c.Tag=$tag;$c.AfterUpdate='=UXFieldChanged("'+$c.Name+'","'+$field.key+'")'
    }else{
        $c=$f.Controls($name);$c.Visible=$true;$c.Tag=$tag;$c.Left=[int]($x);$c.Top=[int]($y+350);$c.Width=[int](4900);$c.Height=[int](360)
        if($field.multiline){$c.Height=[int](2200);$c.Width=[int](10200);$c.ScrollBars=2}
        $c.AfterUpdate='=UXRefreshEditor()'
    }
}
foreach($name in @('cmdSave','cmdClose','lblSaved','lblRecovery')){$f.Controls($name).Visible=$true}
$f.Controls('cmdSave').Left=[int](11100);$f.Controls('cmdSave').Top=[int](7300);$f.Controls('cmdSave').Width=[int](2200);$f.Controls('cmdSave').Height=[int](500);$f.Controls('cmdSave').OnClick='=UXSaveEditor()'
$f.Controls('cmdClose').Left=[int](3100);$f.Controls('cmdClose').Top=[int](7300);$f.Controls('cmdClose').Width=[int](2400);$f.Controls('cmdClose').Height=[int](500);$f.Controls('cmdClose').Caption='Back to reports';$f.Controls('cmdClose').OnClick='=UXNavigate("frmLibrary")'
$f.Controls('lblSaved').Left=[int](3100);$f.Controls('lblSaved').Top=[int](6800);$f.Controls('lblSaved').Width=[int](10000);$f.Controls('lblSaved').Height=[int](350)
$f.Controls('lblRecovery').Left=[int](3100);$f.Controls('lblRecovery').Top=[int](7850);$f.Controls('lblRecovery').Width=[int](10000);$f.Controls('lblRecovery').Height=[int](250)
[void](Control $f 100 'uxDirty' 5900 7350 4900 350 'No pending changes')
[void](Control $f 104 'uxPrevious' 3100 6300 1600 400 'Previous page' '=UXEditorPage(-1)')
[void](Control $f 100 'uxPage' 5000 6350 4500 300 'Page 1')
[void](Control $f 104 'uxNext' 11100 6300 2200 400 'Next page' '=UXEditorPage(1)')
foreach($entry in @(@('cmdRoute','Review routes','=UXEditorReview("Review")',8),@('cmdComments','Review comments','=CommentsFromEditor()',8),@('cmdChecklist','Draft checks','=UXEditorReview("Review")',8),@('cmdCompare','Compare revisions','=CompareSelectedRevision()',9),@('cmdRestore','Restore as new revision','=RestoreSelected()',9),@('cmdSnippets','Insert narrative snippet','=OpenScreen("frmSnippets")',5))){
    $c=$f.Controls($entry[0]);$c.Left=[int](3100);$actionSlot=if($entry[0] -in @('cmdComments','cmdRestore')){1}elseif($entry[0] -eq 'cmdChecklist'){2}else{0};$c.Top=[int](3900+$actionSlot*700);$c.Width=[int](4900);$c.Height=[int](500);$c.Caption=$entry[1];$c.OnClick=$entry[2];$c.Tag='UX:'+$entry[3]+':0:Action'
}
$f.Controls('lstRevisions').Left=[int](8200);$f.Controls('lstRevisions').Top=[int](3900);$f.Controls('lstRevisions').Width=[int](5100);$f.Controls('lstRevisions').Height=[int](2000);$f.Controls('lstRevisions').Tag='UX:9:0:History'
$f.Controls('txtTraitGuidance').Left=[int](3100);$f.Controls('txtTraitGuidance').Top=[int](3900);$f.Controls('txtTraitGuidance').Width=[int](10200);$f.Controls('txtTraitGuidance').Height=[int](2200);$f.Controls('txtTraitGuidance').Tag='UX:4:2:Guidance'
SaveScreen $f 'frmEditor'
NewScreen 'frmUXDecision' 'Keep your changes?' 3700;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true
[void](Control $f 100 'lblBody' 400 1000 12600 1000 'Pending values have not been saved.')
[void](Control $f 104 'cmdCancel' 400 2400 3500 550 'Keep editing' '=UXDecide("Cancel")')
[void](Control $f 104 'cmdDiscard' 4600 2400 3500 550 'Discard changes' '=UXDecide("Discard")')
[void](Control $f 104 'cmdSave' 8800 2400 4000 550 'Save and continue' '=UXDecide("Save")')
$f.OnLoad='=UXLoadDecision()';SaveScreen $f 'frmUXDecision'
NewScreen 'frmUXGroups' 'Summary groups' 7700;$f=$script:swiftForm;UXChrome $f
$f.Controls('lblHeading').Top=[int](1200)
[void](Control $f 100 'lblInfo' 400 1850 12600 500 'Choose a group to manage members, shared report blocks, review and comparison. Folders remain separate.')
[void](Control $f 110 'lstGroups' 400 2500 12600 3700);$c=$script:swiftControl;$c.ColumnCount=5;$c.ColumnHeads=$true;$c.ColumnWidths='0;5600;1700;1300;1000';$c.OnDblClick='=UXOpenGroup()'
[void](Control $f 104 'cmdOpen' 400 6500 3700 550 'Open selected group' '=UXOpenGroup()')
[void](Control $f 104 'cmdNew' 4600 6500 3700 550 'Create a group' '=OpenScreen("frmMajorWorkbench")')
$f.OnLoad='=UXLoadGroups()';$f.OnActivate='=UXLoadGroups()';SaveScreen $f 'frmUXGroups'
foreach($screen in @('frmUXResources','frmUXSettings')){
    NewScreen $screen $(if($screen -eq 'frmUXResources'){'Resources'}else{'Settings & Help'}) 7600;$f=$script:swiftForm;UXChrome $f;$f.Controls('lblHeading').Top=[int](1300)
    $items=if($screen -eq 'frmUXResources'){@(@('Narrative snippets','frmSnippets'),@('Review route templates','frmMajorWorkbench'),@('Reference guidance','frmUXHelp'))}else{@(@('Contributors and My Work','frmMyWork'),@('Backup and restore','frmBackup'),@('Operation history and import undo','frmOperations'),@('Recover pending edits','frmRecovery'),@('Support and limitations','frmSupport'))}
    for($i=0;$i -lt $items.Count;$i++){[void](Control $f 104 ('cmdTool'+$i) 400 (2300+$i*850) 12600 600 $items[$i][0] ('=OpenScreen("'+$items[$i][1]+'")'))}
    SaveScreen $f $screen
}
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome');UXChrome $f;$f.Controls('lblHeading').Caption='Overview';$f.Controls('lblHeading').Top=[int](1100)
foreach($name in @('lblCommand','lblProfile','lblPrototype','lblUse')){$f.Controls($name).Visible=$false}
$f.Controls('cmdMajor').Caption='Summary groups';$f.Controls('cmdMajor').OnClick='=UXNavigate("frmUXGroups")'
$f.Controls('cmdSnippets').Caption='Resources';$f.Controls('cmdSnippets').OnClick='=UXNavigate("frmUXResources")'
$f.Controls('cmdSupport').Caption='Settings & Help';$f.Controls('cmdSupport').OnClick='=UXNavigate("frmUXSettings")'
$f.Section(0).Height=[int](8200)
$visible=@('cmdLibrary','cmdMyWork','cmdMajor','cmdRecovery','cmdSnippets','cmdSupport','cmdImportReview','cmdPDF','cmdBackup','cmdOperations')
for($i=0;$i -lt $visible.Count;$i++){$c=$f.Controls($visible[$i]);$c.Left=[int](400+($i%2)*6500);$c.Top=[int](2100+[math]::Floor($i/2)*1000);$c.Width=[int](6100);$c.Height=[int](650)}
SaveScreen $f 'frmHome'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
foreach($c in $f.Controls){$c.Visible=$false;$c.Tag='WBX:Hidden'}
UXChrome $f;$f.Section(0).Height=[int](12500);$f.ScrollBars=2
[void](Control $f 109 'uxMode' 10 10 100 100);$script:swiftControl.Visible=$false
$modeButtons=@('Members','Defaults','Review','Comparison','History')
for($i=0;$i -lt $modeButtons.Count;$i++){[void](Control $f 104 ('uxMode'+$i) (400+$i*2500) 1200 2350 450 $modeButtons[$i] ('=UXWorkbenchMode("'+$modeButtons[$i]+'")'))}
foreach($name in @('lstReports','txtTitle','cboKind','cboPaygrade','cboGroup','txtNote','lblStatus','lstDetails','txtDetail')){$c=$f.Controls($name);$c.Visible=$true;$c.Tag='WBX:Members:Defaults:Review:Comparison:History'}
$f.Controls('lstReports').Top=[int](1850);$f.Controls('lstReports').Height=[int](2100);$f.Controls('txtTitle').Top=[int](4150);$f.Controls('cboKind').Top=[int](4150);$f.Controls('cboPaygrade').Top=[int](4150);$f.Controls('cboGroup').Top=[int](4750);$f.Controls('txtNote').Top=[int](5350);$f.Controls('txtNote').Height=[int](550)
$f.Controls('lstDetails').Top=[int](8300);$f.Controls('lstDetails').Height=[int](1100);$f.Controls('txtDetail').Top=[int](9500);$f.Controls('txtDetail').Height=[int](1600);$f.Controls('lblStatus').Top=[int](11400);$f.Controls('lblStatus').Height=[int](600)
$actionsByMode=@{Members=@('CreateGroup','AddMembers','RemoveMember');Defaults=@('SaveDefault','SuggestDefault');Review=@('CreateRoute','StartRoute','ShowRoute','AcceptStep','ReturnStep','CancelRoute','RunChecks','ApproveRoute','ReopenApproval');Comparison=@();History=@('ShowBatch','PreviewUndo','ApplyBatch')}
foreach($mode in $actionsByMode.Keys){$index=0;foreach($action in $actionsByMode[$mode]){$c=$f.Controls('cmd'+$action);$c.Tag='WBX:'+$mode;$c.Left=[int](400+($index%4)*3200);$c.Top=[int](6200+[math]::Floor($index/4)*550);$c.Width=[int](3000);$c.Height=[int](450);$index++}}
foreach($pair in @(@('cboTemplate','Review'),@('cboRun','Review'),@('cboReviewer','Review'),@('txtReviewerNames','Review'),@('cmdAddReviewer','Review'),@('cmdClearReviewers','Review'),@('cboSnapshot','History'),@('cboMeeting','History'),@('cboBatch','History'),@('cboField','Defaults'),@('txtValue','Defaults'),@('chkNull','Defaults'),@('cboScope','Defaults'),@('cboFolder','Defaults'))){$c=$f.Controls($pair[0]);$c.Tag='WBX:'+$pair[1];$c.Top=[int](7700);$c.Left=[int](400);$c.Width=[int](6000);$c.Height=[int](400)}
# Per-mode selectors occupy separate rows/columns rather than overlapping.
$modeSelectors=@{Review=@('cboTemplate','cboRun','cboReviewer','txtReviewerNames','cmdAddReviewer','cmdClearReviewers');Comparison=@();Defaults=@('cboField','txtValue','chkNull','cboScope','cboFolder');History=@('cboBatch','cboSnapshot','cboMeeting')}
foreach($mode in $modeSelectors.Keys){$i=0;foreach($name in $modeSelectors[$mode]){$c=$f.Controls($name);$c.Left=[int](400+($i%3)*4300);$c.Top=[int](7400+[math]::Floor($i/3)*450);$c.Width=[int](4000);$c.Height=[int](350);$i++}}
[void](Control $f 104 'uxBulk' 400 6850 6200 500 'Edit shared report blocks' '=UXStartGroupBulk()');$script:swiftControl.Tag='WBX:Members'
[void](Control $f 104 'uxBack' 6800 6850 6200 500 'Back to summary groups' '=UXNavigate("frmUXGroups")');$script:swiftControl.Tag='WBX:Members:Defaults:Review:Comparison:History'
$f.OnLoad='=UXLoadWorkbench()';SaveScreen $f 'frmMajorWorkbench'
NewScreen 'frmUXBulk' 'Shared report blocks' 14200;$f=$script:swiftForm;UXChrome $f;$f.ScrollBars=2;$f.Controls('lblHeading').Top=[int](1200)
[void](Control $f 100 'uxGroupTitle' 400 1800 12600 400 'Summary group')
[void](Control $f 100 'uxSection' 3100 2350 10000 400 'Member & report details')
$bulkCounts=@{}
foreach($field in $uxFields){
    if($field.key -notin $uxShared){continue}
    $section=[int]$field.section;if(-not $bulkCounts.ContainsKey($section)){$bulkCounts[$section]=0};$ordinal=$bulkCounts[$section];$bulkCounts[$section]++
    $page=[math]::Floor($ordinal/8);$slot=$ordinal%8;$x=3100+($slot%2)*5100;$y=2950+[math]::Floor($slot/2)*750;$tag='BLX:'+$section+':'+$page
    [void](Control $f 100 ('uxLabel'+$field.key) $x $y 4900 280 $field.label);$script:swiftControl.Tag=$tag
    $valueType=if($field.checkbox){106}elseif($field.options){111}else{109};[void](Control $f $valueType ('uxValue'+$field.key) $x ($y+300) 4900 360);$c=$script:swiftControl;if($field.checkbox){$c.TripleState=$true}elseif($field.options){$c.RowSourceType='Value List';$c.RowSource=$field.options -join ';';$c.LimitToList=$false};$c.Tag=$tag;$c.AfterUpdate='=UXBulkChanged("'+$field.key+'")'
    if($valueType -in @(109,111)){$c.OnChange='=UXInvalidateBulkPreview()'}
}
$y=2550
foreach($section in @(0,1,2,6)){for($page=0;$page -lt [math]::Ceiling($bulkCounts[$section]/8);$page++){[void](Control $f 104 ('uxSection'+$section+'Page'+$page) 400 $y 2450 450 ($uxSections[$section]+' '+($page+1)) ('=UXBulkSection('+$section+','+$page+')'));$y+=500}}
[void](Control $f 100 'uxStatus' 3100 6250 10000 600 'Only edited blocks are staged. Mixed values start blank; clearing stages an explicit clear.')
[void](Control $f 100 'lblReason' 400 7000 2500 350 'Reason for changes')
[void](Control $f 109 'txtReason' 3100 6950 10000 450)
$script:swiftControl.AfterUpdate='=UXInvalidateBulkPreview()'
$script:swiftControl.OnChange='=UXInvalidateBulkPreview()'
[void](Control $f 104 'cmdPreview' 3100 7600 4600 550 'Preview block changes' '=UXPreviewBulk()')
[void](Control $f 104 'cmdApply' 8400 7600 4900 550 'Apply reviewed changes' '=UXApplyBulk()')
$script:swiftControl.Enabled=$false
[void](Control $f 110 'lstPreview' 400 8450 12600 1700);$c=$script:swiftControl;$c.ColumnCount=5;$c.ColumnWidths='0;2600;3000;3500;3500';$c.OnClick='=UXBulkDetail()'
[void](Control $f 109 'txtFullValues' 400 10350 12600 2400);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 104 'cmdBack' 400 13100 12600 550 'Return to group' '=OpenScreen("frmMajorWorkbench")')
SaveScreen $f 'frmUXBulk'

$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
[void](Control $f 104 'uxMembers' 6800 6200 6000 500 'Choose group members' '=UXChooseMembers()');$script:swiftControl.Tag='WBX:Members'
[void](Control $f 104 'uxComparison' 6800 6750 6000 500 'Open editable trait comparison' '=UXOpenComparison()');$script:swiftControl.Tag='WBX:Comparison'
SaveScreen $f 'frmMajorWorkbench'
NewScreen 'frmUXMembers' 'Choose group members' 7600;$f=$script:swiftForm
[void](Control $f 100 'lblInfo' 400 1000 12600 700 'Choose the explicit roster. Verify paygrade and eligibility; profile attribution is local. Unavailable/archived existing members require separate deliberate removal.')
[void](Control $f 110 'lstReports' 400 1900 12600 3000);$c=$script:swiftControl;$c.ColumnCount=4;$c.ColumnWidths='0;5000;2000;5000';$c.MultiSelect=2
[void](Control $f 100 'lblNote' 400 5100 12600 350 'Membership / exception note')
[void](Control $f 109 'txtNote' 400 5500 12600 550)
[void](Control $f 104 'cmdSave' 400 6350 6100 550 'Save explicit membership' '=UXSaveMembers()')
[void](Control $f 104 'cmdCancel' 6900 6350 6100 550 'Cancel' '=UXCloseScreen("frmUXMembers")')
[void](Control $f 100 'lblStatus' 400 7050 12600 450 '')
SaveScreen $f 'frmUXMembers'
NewScreen 'frmUXComparison' 'Trait comparison' 12200;$f=$script:swiftForm;UXChrome $f;$f.ScrollBars=2;$f.Controls('lblHeading').Top=1200
[void](Control $f 109 'txtMeeting' 10 10 100 100);$script:swiftControl.Visible=$false
[void](Control $f 110 'lstTraits' 400 2000 12600 2700);$c=$script:swiftControl;$c.ColumnCount=12;$c.ColumnHeads=$true;$c.ColumnWidths='0;2400;750;750;750;750;750;750;750;900;2100;0';$c.OnClick='=UXSelectComparisonMember()'
[void](Control $f 100 'uxMember' 400 4900 12600 400 'Select a member above to edit proposed marks in this comparison window.')
for($block=33;$block -le 39;$block++){
    $x=400+($block-33)*1800
    [void](Control $f 100 ('lblTrait'+$block) $x 5450 1600 300 ('Block '+$block))
    [void](Control $f 111 ('cboTrait'+$block) $x 5850 1600 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='NOB;1.0;2.0;3.0;4.0;5.0';$c.LimitToList=$false;$c.AfterUpdate='=UXComparisonAverage()'
}
[void](Control $f 100 'uxITA' 400 6450 12600 400 'Proposed ITA: no member selected')
[void](Control $f 111 'cboPromotion' 400 7000 6100 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='NOB;Significant Problems;Progressing;Promotable;Must Promote;Early Promote';$c.LimitToList=$false
[void](Control $f 104 'cmdStage' 6900 7000 6100 500 'Stage selected member changes' '=UXStageComparison()')
[void](Control $f 110 'lstProposals' 400 7750 12600 2100);$c=$script:swiftControl;$c.ColumnCount=5;$c.ColumnHeads=$true;$c.ColumnWidths='0;3000;3300;3300;2500';$c.MultiSelect=2
[void](Control $f 104 'cmdApply' 400 10200 6100 550 'Apply reviewed selected proposals' '=UXApplyComparison()')
[void](Control $f 104 'cmdBack' 6900 10200 6100 550 'Return to group' '=OpenScreen("frmMajorWorkbench")')
[void](Control $f 100 'lblStatus' 400 11000 12600 850 'Saved-value baseline captured automatically. Comments are not shown. Pending edits are proposals until applied; stale groups and locked reports reject changes.')
$f.OnLoad='=UXLoadComparison()';SaveScreen $f 'frmUXComparison'
