. (Join-Path $PSScriptRoot 'access-properties.ps1')
# Final additive UI pass: table history, dedicated resources, explicit live-group tools.
$db=$app.CurrentDb();if($db.TableDefs('RSCABaseline').Fields('Paygrade').Size -lt 3){$db.Execute('ALTER TABLE RSCABaseline ALTER COLUMN Paygrade TEXT(5)',128)}
if(-not $script:swiftManagedBuild){foreach($m in @('modUXSeniorTable','modUXResourceTools','modRSDirectory','modRSDirectoryUI','modUXNAVFIT','modUXTraitGrid','modUXGroupPages','modUXViewport','modUXDashboard','modUXHelp','modUXBlockTools')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
function AddResourceControl($f,$type,$name,$x,$y,$w,$h,$caption='',$event=''){
 $c=$app.CreateControl($f.Name,$type,0,'','',[int]$x,[int]$y,[int]$w,[int]$h);$c.Name=$name
 if($type -in @(100,104)){$c.Caption=$caption}
 if($type -ne 112){$c.FontName='Segoe UI';$c.FontSize=10}
 if($event){$c.OnClick=$event}
 if($type -in @(104,109,110,111)){$c.BackColor=16777215;$c.ForeColor=2368548;$c.BorderColor=14012862}
 $script:resourceControl=$c
}
$app.LoadFromText(2,'frmUXRSCATable',(Join-Path $root 'src/BlankForm.txt'))
$app.DoCmd.OpenForm('frmUXRSCATable',1);$f=$app.Forms.Item('frmUXRSCATable')
$f.Width=10100;$f.RecordSelectors=$false;$f.NavigationButtons=$false;$f.ScrollBars=2;$f.Section(0).BackColor=16777215
$headers=@(@('grade','Paygrade',100,800),@('avg','RSCA',1000,1600),@('count','Previous graded traits',2900,2600),@('date','Date verified / pulled from BOL',5800,3700))
foreach($h in $headers){AddResourceControl $f 100 ('head'+$h[0]) $h[2] 50 $h[3] 400 $h[1]}
$grades=@('E1','E2','E3','E4','E5','E6','E7','E8','E9','W1','W2','W3','W4','W5','O1','O2','O3','O4','O5','O6','O7','O8','O9','O10')
$i=0
foreach($grade in $grades){
 $y=500+($i*390);AddResourceControl $f 100 ('grade'+$grade) 100 $y 800 330 $grade
 foreach($h in $headers[1..3]){
  AddResourceControl $f 109 ($h[0]+$grade) $h[2] $y $h[3] 330
  $c=$script:resourceControl;$c.AfterUpdate='=RSTableChanged()';$c.ControlTipText=$grade+' - '+$h[1];$c.TabIndex=($i*3+[array]::IndexOf(@('avg','count','date'),$h[0]))
  if($h[0] -eq 'date'){$c.Format='yyyy-mm-dd';$c.ShowDatePicker=1}else{$c.Format=if($h[0] -eq 'avg'){'0.00'}else{'0'}}
 }
 $i++
}
$f.Section(0).Height=10000;$f.OnResize='=RSTableResize()'
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups')
foreach($name in @('uxExportSelected','uxPDFSelected')){if(@($f.Controls|ForEach-Object {$_.Name}) -contains $name){$app.DeleteControl($f.Name,$name)}}
AddResourceControl $f 104 'uxExportSelected' 2850 9800 5050 450 'Export selected group to NAVFIT98A' '=UXExportSelectedGroup()'
AddResourceControl $f 104 'uxPDFSelected' 8150 9800 5150 450 'Preview / create selected group PDFs' '=PDFOutputOpenGroup("frmUXGroups")'
SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,10450))
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmUXSeniorEditor',1);$f=$app.Forms.Item('frmUXSeniorEditor')
foreach($name in @('uxRSCATable','cmdSaveRSTable','cmdRSHistory')){try{$app.DeleteControl($f.Name,$name)}catch{}}
# Keep old single-observation buffers hidden for backend compatibility only.
foreach($c in $f.Controls){if($c.Tag -eq 'SR:3'){$c.Tag='';$c.Visible=$false}}
AddResourceControl $f 112 'uxRSCATable' 2850 3000 10450 3500;$script:resourceControl.SourceObject='Form.frmUXRSCATable';$script:resourceControl.Tag='SR:3'
AddResourceControl $f 104 'cmdSaveRSTable' 2850 6600 5000 450 'Save RSCA table' '=RSSaveTableButton()';$script:resourceControl.Tag='SR:3'
AddResourceControl $f 104 'cmdRSHistory' 8100 6600 5200 450 'View prior baselines' '=RSViewTableHistory()';$script:resourceControl.Tag='SR:3'
$f.Controls('lstBaselines').Left=2850;$f.Controls('lstBaselines').Top=3000;$f.Controls('lstBaselines').Width=10450;$f.Controls('lstBaselines').Height=3500
$f.Controls('lstBaselines').ColumnWidths='900;1600;1200;1800;1400;3300'
$f.OnLoad='=RSLoadSeniorEditor()'
$app.DoCmd.Close(2,$f.Name,1)
# Clone only the shared workspace chrome; template content is dedicated.
$chrome=Join-Path $root 'build/route-resources-chrome.txt';$app.SaveAsText(2,'frmUXResources',$chrome);$app.LoadFromText(2,'frmUXRouteTemplates',$chrome)
$app.DoCmd.OpenForm('frmUXRouteTemplates',1);$f=$app.Forms.Item('frmUXRouteTemplates')
foreach($name in @($f.Controls|Where-Object {$_.Name -notmatch '^ux(Nav\d|Rail$|Brand$|NavCaption$|Context$|CompactNav$)' -and $_.Name -ne 'lblHeading'}|ForEach-Object {$_.Name})){$app.DeleteControl($f.Name,$name)}
$f.Caption='Review route templates';$f.Controls('lblHeading').Caption='Review route templates'
AddResourceControl $f 110 'lstTemplates' 2850 1800 3600 2200;$c=$script:resourceControl;$c.ColumnCount=2;$c.ColumnWidths='0;3200';$c.AfterUpdate='=UXShowTemplate()'
AddResourceControl $f 110 'lstSteps' 6650 1800 6650 2200;$c=$script:resourceControl;$c.ColumnCount=3;$c.ColumnWidths='600;2000;3700';$c.ColumnHeads=$true
AddResourceControl $f 100 'lblNew' 2850 4150 10450 350 'Create a new reusable route (reviewers run in the order added)'
AddResourceControl $f 100 'lblTitle' 2850 4550 1800 300 'Template name'
AddResourceControl $f 109 'txtTitle' 4750 4500 8550 400
AddResourceControl $f 111 'cboReviewer' 2850 5050 4600 400;$c=$script:resourceControl;$c.ColumnCount=2;$c.ColumnWidths='0;4000';$c.LimitToList=$true
AddResourceControl $f 104 'cmdAdd' 7700 5050 2600 400 'Add reviewer' '=UXTemplateReviewer(False)'
AddResourceControl $f 104 'cmdClear' 10600 5050 2700 400 'Clear reviewer order' '=UXTemplateReviewer(True)'
AddResourceControl $f 109 'txtIDs' 10 10 100 100;$script:resourceControl.Visible=$false
AddResourceControl $f 109 'txtReviewers' 2850 5600 10450 550;$script:resourceControl.Locked=$true
AddResourceControl $f 100 'lblInstructions' 2850 6250 10450 300 'Reviewer instructions'
AddResourceControl $f 109 'txtInstructions' 2850 6600 10450 750
AddResourceControl $f 104 'cmdSave' 8200 7500 5100 450 'Create route template' '=UXSaveTemplate()'
AddResourceControl $f 104 'cmdBack' 2850 7500 5000 450 'Back to Resources' '=UXCloseTemplates()'
AddResourceControl $f 100 'lblStatus' 2850 8100 10450 600 'Select a template to inspect its ordered steps, or create a new one.'
$f.OnLoad='=UXLoadRouteTemplates()';$f.Section(0).Height=8900
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmUXResources',1);$f=$app.Forms.Item('frmUXResources');$f.Controls('cmdTool1').OnClick='=UXOpenRouteTemplates()';$app.DoCmd.Close(2,$f.Name,1)
# Keep group actions below content; Members is the sole live comparison tab.
foreach($screen in @('frmUXComparison','frmUXBulk','frmMajorWorkbench')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 $already=$false;try{$null=$f.Controls('uxCurrentExport');$already=$true}catch{}
 if($already -and $f.Controls('uxCurrentExport').Tag -ne 'GroupFooter'){
  foreach($c in $f.Controls){if($c.Left -ge 2450 -and $c.Top -ge 3650){$c.Top=[int]($c.Top-1050)}}
 }
 foreach($name in @('uxSeniorCaption','uxCurrentSenior','uxCurrentExport')){try{$app.DeleteControl($f.Name,$name)}catch{}}
 $bottom=2600;foreach($c in $f.Controls){if($c.Visible -and $c.Left -ge 2450){$bottom=[math]::Max($bottom,$c.Top+$c.Height)}}
 AddResourceControl $f 100 'uxSeniorCaption' 2850 ($bottom+150) 10450 300 'Reporting senior';$script:resourceControl.Tag='GroupFooter'
 AddResourceControl $f 104 'uxCurrentSenior' 2850 ($bottom+500) 5100 430 'Reporting senior / PSG RSCA' ('=UXCurrentGroupTool("'+$screen+'","Senior")');$script:resourceControl.Tag='GroupFooter'
 AddResourceControl $f 104 'uxCurrentExport' 8200 ($bottom+500) 5100 430 'Export group to NAVFIT98A' ('=UXCurrentGroupTool("'+$screen+'","Export")');$script:resourceControl.Tag='GroupFooter'
 try{$app.DeleteControl($screen,'uxGroupTab3')}catch{}
 $tabs=@(0,1,2,4);$lefts=@(2850,4900,9100,11200);$widths=@(1900,4050,1950,2100)
 for($i=0;$i -lt 4;$i++){$c=$f.Controls('uxGroupTab'+$tabs[$i]);$c.Left=$lefts[$i];$c.Width=$widths[$i]}
 SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,[int]($bottom+1100)))
 if($screen -eq 'frmUXComparison'){
  AddResourceControl $f 100 'uxGroupMeta' 2850 1500 6800 350 ''
  $f.Caption='Summary group';$f.Controls('uxTraitGrid').Top=3000;$f.Controls('uxTraitGrid').Height=2900
  $f.Controls('uxGridAverage').Top=6050;$f.Controls('uxGridAverage').Height=350
  $actions=@(@('uxMembers',2850,6500,3300,'Choose group members','Members'),@('uxBulk',6350,6500,3300,'Change blocks on existing reports','Blocks'),@('uxPolicyReview',9850,6500,3450,'Grouping policy and roster review','Policy'),@('uxOpenGroupReport',6350,7150,3300,'Open selected report','Open'),@('uxValidation',9850,7150,3450,'All summary-group validation issues','Validation'))
  foreach($a in $actions){AddResourceControl $f 104 $a[0] $a[1] $a[2] $a[3] 500 $a[4] ('=UXMemberAction("'+$a[5]+'")');$script:resourceControl.Tag='GroupMemberAction'}
  $f.Controls('cmdBack').Left=2850;$f.Controls('cmdBack').Top=7150;$f.Controls('cmdBack').Width=3300
  $f.Controls('lblStatus').Top=7850;$f.Controls('lblStatus').Height=500
  $f.Controls('cmdRetryCells').Top=8450;$f.Controls('cmdDiscardCells').Top=8450
  $f.Controls('uxSeniorCaption').Top=9100;$f.Controls('uxCurrentSenior').Top=9450;$f.Controls('uxCurrentExport').Top=9450
  $f.Section(0).Height=10200
 }
 $app.DoCmd.Close(2,$f.Name,1)
}

$app.DoCmd.OpenForm('frmUXGroupRSCA',1);$f=$app.Forms.Item('frmUXGroupRSCA')
try{$app.DeleteControl($f.Name,'cmdViewSenior')}catch{}
AddResourceControl $f 104 'cmdViewSenior' 2850 8800 10450 450 'Open selected reporting senior record' '=RSViewGroupSenior()'
SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,[int]9500))
$app.DoCmd.Close(2,$f.Name,1)
