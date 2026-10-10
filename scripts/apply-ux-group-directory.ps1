# Schema 25 is required to retain all 35 RUAD columns without changing existing report fields.
$profileDB=$app.CurrentDb()
if(-not @($profileDB.TableDefs.Item('CommandMember').Fields|ForEach-Object {$_.Name}).Contains('RUADData')){$profileDB.Execute('ALTER TABLE CommandMember ADD COLUMN RUADData LONGTEXT',128)}
$profileDB=$app.CurrentDb();$profileDB.TableDefs.Refresh();$profileDB.TableDefs.Item('CommandMember').Fields.Refresh();$profileDB.TableDefs.Item('CommandMember').Fields.Item('RUADData').AllowZeroLength=$true
$profileDB.Execute('UPDATE AppMetadata SET SchemaVersion=25',128)
foreach($moduleName in @('modRUADProfiles','modCommandDirectory','modUXDirectory','modUXGroupPages','modDirectoryBatch','modUXDashboard','modUXViewport','modUX','modUXTraitGrid','modUXBlockTools')){$app.LoadFromText(5,$moduleName,(Join-Path $root ('src/'+$moduleName+'.bas')))}

# Compact, complete profile. All pending field values remain in unbound controls across page changes.
$app.DoCmd.DeleteObject(2,'frmUXMemberEditor')
NewScreen 'frmUXMemberEditor' 'Command member profile' 8500;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.Width=12200;$f.ScrollBars=2
$f.Controls('lblHeading').Width=11200
$items=@(@('txtName','NAME / Member display name',400,5500),@('txtReportName','Official report name (27 characters)',6200,5500),@('txtRank','Report rank/rate (5 characters)',400,2700),@('txtSSN','SSN (nine digits, or unentered)',3450,2700))
for($i=0;$i -lt $items.Count;$i++){$v=$items[$i];$y=if($i -lt 2){1000}else{1900};[void](Control $f 100 ('lbl'+$v[0]) $v[2] $y $v[3] 320 $v[1]);[void](Control $f 109 $v[0] $v[2] ($y+350) $v[3] 420);$script:swiftControl.TabIndex=$i}
$pages=@('Unit','Member','Assignment','Mobilization','Contact')
for($i=0;$i -lt 5;$i++){[void](Control $f 104 ('ruadPage'+$i) (400+2300*$i) 2850 2150 450 $pages[$i] ('=UXRUADPage('+ $i +')'))}
$fields=Get-Content (Join-Path $root 'src/ruad-profile-fields.json') -Raw|ConvertFrom-Json
$slots=@(0,0,0,0,0)
for($i=0;$i -lt $fields.Count;$i++){
 $item=$fields[$i];if($item.key -eq 'NAME'){continue};$slot=$slots[$item.page];$slots[$item.page]++
 $x=400+3850*($slot%3);$y=3500+1050*[int][Math]::Floor($slot/3)
 [void](Control $f 100 ('ruadLabel'+$i) $x $y 3600 400 $item.key)
 [void](Control $f 109 ('ruad'+$i) $x ($y+440) 3600 450);$c=$script:swiftControl
 $c.ControlTipText=$item.key+' - retained as RUAD text; identifiers keep leading zeroes'
 if($item.key -eq 'PRD'){$c.Format='Short Date';$c.ShowDatePicker=1}
}
[void](Control $f 100 'lblStatus' 400 6850 11400 500 'All 35 RUAD columns are available. NAME is maintained above; RUAD grade and report rank/rate are separate. Profile edits do not rewrite reports.')
[void](Control $f 104 'cmdSave' 400 7500 5400 500 'Save member profile' '=UXDirectorySave()')
[void](Control $f 104 'cmdCancel' 6300 7500 5400 500 'Cancel' '=UXDirectoryClose("frmUXMemberEditor")')
$tab=0;foreach($name in @('txtName','txtReportName','txtRank','txtSSN','ruadPage0','ruadPage1','ruadPage2','ruadPage3','ruadPage4')){$f.Controls($name).TabIndex=[int]$tab;$tab++}
for($i=0;$i -lt $fields.Count;$i++){if($fields[$i].key -ne 'NAME'){$f.Controls('ruad'+$i).TabIndex=[int]$tab;$tab++}}
$f.Controls('cmdSave').TabIndex=[int]$tab;$f.Controls('cmdCancel').TabIndex=[int]($tab+1)
SaveScreen $f 'frmUXMemberEditor'

# Group report creation is explicit; cancelling does not create an empty draft.
NewScreen 'frmUXGroupReport' 'Add report to summary group' 7200;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.Width=11800
$f.Controls('lblHeading').Width=11000
[void](Control $f 100 'lblGroup' 400 1000 11000 500)
[void](Control $f 100 'lblMember' 400 1650 11000 300 'Choose an existing command member, or enter a new member below')
[void](Control $f 111 'cboMember' 400 2050 11000 450);$c=$script:swiftControl;$c.ColumnCount=6;$c.ColumnWidths='0;6000;1500;0;0;0';$c.LimitToList=$true;$c.AfterUpdate='=UXGroupReportMember()'
$items=@(@('txtName','Member display name'),@('txtOfficial','Official report name (27 characters)'),@('txtRank','Rank/rate matching group paygrade'),@('txtSSN','SSN (nine digits, or unentered)'),@('txtFrom','14. From'),@('txtTo','15. To'))
for($i=0;$i -lt $items.Count;$i++){$x=400+5700*($i%2);$y=2750+900*[int][Math]::Floor($i/2);$v=$items[$i];[void](Control $f 100 ('lbl'+$v[0]) $x $y 5300 320 $v[1]);[void](Control $f 109 $v[0] $x ($y+350) 5300 420);if($i -ge 4){$script:swiftControl.Format='Short Date';$script:swiftControl.ShowDatePicker=1}}
[void](Control $f 109 'txtMemberVersion' 10 10 100 100);$script:swiftControl.Visible=$false
[void](Control $f 100 'lblStatus' 400 5600 11000 550 'Uses the group report type and starting defaults. Create saves the draft and adds it to this group. Dates can remain unentered while drafting.')
[void](Control $f 104 'cmdCreate' 400 6400 5300 500 'Create report' '=UXGroupReportCreate()')
[void](Control $f 104 'cmdCancel' 6100 6400 5300 500 'Cancel' '=UXDirectoryClose("frmUXGroupReport")')
SaveScreen $f 'frmUXGroupReport'

$tabLeft=@(2850,4550,7650,9350,11400);$tabWidth=@(1600,3000,1600,1950,1900)
$captions=@('Members','Defaults for new reports','Review','Comparison','History')
foreach($screen in @('frmMajorWorkbench','frmUXComparison','frmUXBulk')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 # Space is reserved once in design, so normal navigation uses cached responsive geometry.
 $newBar=-not @($f.Controls|ForEach-Object {$_.Name}).Contains('uxGroupContext')
 if($newBar -and $screen -ne 'frmMajorWorkbench'){foreach($c in $f.Controls){if($c.Left -ge 2450 -and $c.Top -ge 1600 -and $c.Name -notmatch '^uxMode[0-4]$'){$c.Top=[int]($c.Top+900)}}}
 if($newBar){[void](Control $f 109 'uxGroupContext' 10 10 100 100);$script:swiftControl.Visible=$false}
 if($newBar){for($i=0;$i -lt 5;$i++){$event='=UXGroupNavigate("'+$screen+'",'+$i+')';[void](Control $f 104 ('uxGroupTab'+$i) $tabLeft[$i] 2050 $tabWidth[$i] 450 $captions[$i] $event)}}
 if($newBar){$event='=UXGroupAddReport("'+$screen+'")';[void](Control $f 104 'uxAddGroupReport' 10800 1000 2500 450 'Add report' $event)}
 if($screen -eq 'frmMajorWorkbench'){$f.Controls('uxGroupTitle').Width=7600;$f.Controls('lblStatus').Top=8300;$f.Controls('lblStatus').Height=280;for($i=0;$i -lt 5;$i++){$f.Controls('uxMode'+$i).Visible=$false}}
 if($screen -eq 'frmUXBulk'){$f.Controls('lblHeading').Width=7600;for($i=0;$i -lt 5;$i++){$f.Controls('uxGroupTab'+$i).Properties.Item('FontBold').Value=($i -eq 1)}}
 if($screen -eq 'frmUXComparison'){$f.Controls('lblHeading').Width=7600;$f.Controls('uxGroupTab0').Properties.Item('FontBold').Value=$true}
 if($newBar){$f.Section(0).Height=[int]($f.Section(0).Height+900)}
 SaveScreen $f $screen
}
# Overview: separate header contributor, metrics, recovery actions and recent reports.
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome')
foreach($entry in @(@('uxContributorLabel',7800,850,5500,300),@('uxContributor',7800,1200,5500,420),@('uxContributorStatus',7800,1640,5500,300),@('uxNoteRecovery',10500,4120,2800,500),@('uxRecoveryStatus',2850,3550,7300,500),@('cmdRecovery',10500,3550,2800,500),@('uxRecentCaption',2850,4770,10450,400),@('uxRecentReports',2850,5270,10450,2300),@('uxOpenRecent',2850,7870,3000,500),@('uxNewReport',9850,7870,3450,500))){$c=$f.Controls.Item($entry[0]);$c.Left=[int]$entry[1];$c.Top=[int]$entry[2];$c.Width=[int]$entry[3];$c.Height=[int]$entry[4]}
foreach($name in @('cmdLibrary','cmdMajor','cmdMyWork')){$f.Controls($name).Top=2200}
foreach($name in @('uxReportCount','uxGroupCount','uxWorkCount')){$f.Controls($name).Top=2850}
$f.Section(0).Height=8800
SaveScreen $f 'frmHome'

# Apply the existing Siltstone palette only to newly created controls/forms.
foreach($screen in @('frmUXMemberEditor','frmUXGroupReport','frmMajorWorkbench','frmUXComparison','frmUXBulk')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 if($screen -in @('frmUXMemberEditor','frmUXGroupReport')){$f.Section(0).BackColor=15397112}
 foreach($c in $f.Controls){
  if($screen -notin @('frmUXMemberEditor','frmUXGroupReport') -and $c.Name -notmatch '^uxGroupTab|^uxAddGroupReport'){continue}
  if($c.ControlType -eq 100){$c.ForeColor=2501670}
  if($c.ControlType -eq 104){$c.UseTheme=$false;$c.BackColor=7254527;$c.ForeColor=2501670;$c.BorderColor=5948108;if($c.Name -in @('cmdSave','cmdCreate')){$c.BackColor=5948108}}
  if($c.ControlType -in @(109,111)){$c.BackColor=16777215;$c.ForeColor=2501670;$c.BorderColor=9682379}
 }
 SaveScreen $f $screen
}
