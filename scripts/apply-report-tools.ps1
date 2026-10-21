# Additive schema and genuine embedded report tools. Earlier records remain intact.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
$db=$app.CurrentDb();$tables=@($db.TableDefs|ForEach-Object {$_.Name})
if($tables -notcontains 'ReportMessage'){
 $db.Execute('CREATE TABLE ReportMessage (MessageID COUNTER CONSTRAINT PK_ReportMessage PRIMARY KEY,MessageGuid TEXT(38) NOT NULL,ReportID LONG NOT NULL,ParentMessageID LONG,AuthorProfileID LONG,ContributorName TEXT(120),CreatedAt DATETIME,MessageBody LONGTEXT,LegacyNoteID LONG,LegacyCommentID LONG,LegacyRevisionID LONG)',128)
 foreach($sql in @('CREATE UNIQUE INDEX UX_MessageGuid ON ReportMessage (MessageGuid)','ALTER TABLE ReportMessage ADD CONSTRAINT FK_MessageReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)','ALTER TABLE ReportMessage ADD CONSTRAINT FK_MessageParent FOREIGN KEY (ParentMessageID) REFERENCES ReportMessage (MessageID)','ALTER TABLE ReportMessage ADD CONSTRAINT FK_MessageAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)','ALTER TABLE ReportMessage ADD CONSTRAINT FK_MessageNote FOREIGN KEY (LegacyNoteID) REFERENCES ReportNote (NoteID)','ALTER TABLE ReportMessage ADD CONSTRAINT FK_MessageComment FOREIGN KEY (LegacyCommentID) REFERENCES ReviewComment (CommentID)','ALTER TABLE ReportMessage ADD CONSTRAINT FK_MessageRevision FOREIGN KEY (LegacyRevisionID) REFERENCES ReportRevision (RevisionID)')){$db.Execute($sql,128)}
 $db.TableDefs.Refresh()
 foreach($field in @('LegacyNoteID','LegacyCommentID')){$table=$db.TableDefs.Item('ReportMessage');$index=$table.CreateIndex('UX_Message_'+$field);$index.Unique=$true;$index.IgnoreNulls=$true;$index.Fields.Append($index.CreateField($field));$table.Indexes.Append($index)}
}
$fields=@($db.TableDefs.Item('ReportNoteDraft').Fields|ForEach-Object {$_.Name})
if($fields -notcontains 'ParentMessageID'){$db.Execute('ALTER TABLE ReportNoteDraft ADD COLUMN ParentMessageID LONG',128);$db.Execute('ALTER TABLE ReportNoteDraft ADD CONSTRAINT FK_NoteDraftMessage FOREIGN KEY (ParentMessageID) REFERENCES ReportMessage (MessageID)',128)}
$reportToolsVersion=[regex]::Match((Get-Content (Join-Path $root 'src/modSwift.bas') -Raw),'Public Const APP_VERSION As String = "([^"]+)"').Groups[1].Value
if(-not $reportToolsVersion){throw 'Missing canonical app version'}
$db.Execute("UPDATE AppMetadata SET SchemaVersion=26,AppVersion='$reportToolsVersion'",128)
$db.Close()
if(-not $script:swiftManagedBuild){foreach($m in @('modReportMessages','modReportTools','modSwift','modMigration','modWorkspaceImport','modReportPages','modUX','modUXReview','modUXReviewLayout','modUXHeaderReview','modReviewComments','modUXNotes','modRouting','modUXMyWork','modUXReadRows','modUXDashboard','modUXValidation','modChecklist','modUXCalculationContext','modUXHelp')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
function ToolRect($form,[string]$name,[int]$x,[int]$y,[int]$w,[int]$h){$c=$form.Controls.Item($name);foreach($v in @(@('Left',$x),@('Top',$y),@('Width',$w),@('Height',$h))){SetAccessProperty $c $v[0] $v[1]}}
function ToolControl($form,[int]$type,[string]$name,[int]$x,[int]$y,[int]$w,[int]$h,[string]$caption='',[string]$event='',[string]$tag=''){
 $c=$app.CreateControl($form.Name,$type,0,'','',$x,$y,$w,$h);$c.Name=$name;$c.FontName='Segoe UI';$c.FontSize=10;$c.Tag=$tag
 if($type -in @(100,104)){$c.Caption=$caption}
 if($event){$c.OnClick=$event}
 if($type -eq 104){$c.UseTheme=$false;$c.BackColor=$app.Eval('UXColor("Neutral")');$c.ForeColor=$app.Eval('UXColor("Ink")')}
 if($type -eq 109){$c.ScrollBars=2;$c.TextFormat=0}
 return $c
}
function ToolForm([string]$name,[int]$height){
 $app.LoadFromText(2,$name,(Join-Path $root 'src/BlankForm.txt'));$app.DoCmd.OpenForm($name,1);$form=$app.Forms.Item($name)
 $form.Width=12400;$form.Section(0).Height=$height;$form.NavigationButtons=$false;$form.RecordSelectors=$false;$form.ScrollBars=2;$form.DividingLines=$false;$form.PopUp=$false;$form.Modal=$false;$form.BorderStyle=0;$form.OnLoad='';$form.OnResize='';$form.Section(0).BackColor=$app.Eval('UXColor("Canvas")')
 return ,$form
}
# Proven route controls become a subform, not an independently opened window.
$all=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
if($all -contains 'frmUXReview'){$app.DoCmd.CopyObject('','frmReportReviewPanel',2,'frmUXReview')}
$app.DoCmd.OpenForm('frmReportReviewPanel',1);$f=$app.Forms.Item('frmReportReviewPanel')
$keep=@('txtReportID','txtGroupID','txtExpectedRevision','txtExpectedStamp','lblHeading','lblSaved','lblTemplate','cboTemplate','cmdStart','lblRun','cboRun','lblStatus','lstSteps','txtStep','lblNote','txtNote','cmdAccept','cmdReturn','cmdRestart','cmdCancel','cmdApprove','cmdReopen','lblUnavailable','cmdBack','cmdRefresh')
foreach($name in @($f.Controls|ForEach-Object {$_.Name})){if($keep -notcontains $name){$app.DeleteControl($f.Name,$name)}}
$f.OnLoad='';$f.OnResize='';$f.OnActivate='';$f.OnTimer='';$f.TimerInterval=0;$f.NavigationButtons=$false;$f.RecordSelectors=$false;$f.DividingLines=$false;$f.PopUp=$false;$f.Modal=$false;$f.BorderStyle=0;$f.ScrollBars=2;$f.Width=12400
foreach($c in $f.Controls){$c.Tag='ROUTE'}
foreach($name in @('txtReportID','txtGroupID','txtExpectedRevision','txtExpectedStamp','lblHeading')){$f.Controls.Item($name).Visible=$false;$f.Controls.Item($name).Tag=''}
foreach($v in @(@('lblSaved',200,650,11800,600),@('lblTemplate',200,1350,5300,300),@('cboTemplate',200,1700,5300,420),@('cmdStart',5700,1700,6300,420),@('lblRun',200,2300,11800,300),@('cboRun',200,2650,11800,420),@('lblStatus',200,3250,11800,900),@('lstSteps',200,4300,5800,1500),@('txtStep',6200,4300,5800,1500),@('lblNote',200,6000,11800,300),@('txtNote',200,6350,11800,850),@('lblUnavailable',200,8650,11800,750),@('cmdBack',200,9550,5700,420),@('cmdRefresh',6200,9550,5800,420))){ToolRect $f $v[0] $v[1] $v[2] $v[3] $v[4]}
$i=0;foreach($name in @('cmdAccept','cmdReturn','cmdRestart','cmdCancel','cmdApprove','cmdReopen')){ToolRect $f $name (200+($i%3)*4000) (7400+[int][Math]::Floor($i/3)*550) 3800 420;$i++}
$f.Controls.Item('lstSteps').ColumnWidths='650;3000;1800;0;0'
$f.Controls.Item('cmdBack').Caption='Back to My Work';$f.Controls.Item('cmdBack').OnClick='=UXWorkBack()';$f.Controls.Item('cmdBack').Tag=''
$c=ToolControl $f 111 'cboReviewMode' 200 100 11800 420;$c.RowSourceType='Value List';$c.RowSource='Route steps;Assigned reviews';$c.DefaultValue='"Route steps"';$c.LimitToList=$true;$c.AfterUpdate='=UXReportReviewMode()'
[void](ToolControl $f 100 'lblAssignments' 200 750 11800 500 'Assigned reviews - use saved revisions; only the assigned reviewer can accept or return.' '' 'ASSIGNMENT')
$c=ToolControl $f 110 'lstAssignments' 200 1350 11800 1800 '' '' 'ASSIGNMENT';$c.ColumnCount=4;$c.ColumnHeads=$true;$c.ColumnWidths='0;5000;3300;3300';$c.AfterUpdate='=UXReportAssignmentSelect()'
$c=ToolControl $f 109 'txtAssignmentHistory' 200 3350 11800 1700 '' '' 'ASSIGNMENT';$c.Locked=$true
[void](ToolControl $f 100 'lblAssignmentNote' 200 5250 11800 350 'Decision note - required when returning; separate from shared report notes' '' 'ASSIGNMENT')
$c=ToolControl $f 109 'txtAssignmentNote' 200 5700 11800 1300 '' '' 'ASSIGNMENT';$c.EnterKeyBehavior=$true
$i=0;foreach($pair in @(@('Accept','Accept assignment'),@('Return','Return assignment'),@('Cancel','Cancel assignment'))){[void](ToolControl $f 104 ('cmdAssignment'+$pair[0]) (200+$i*4000) 7350 3800 450 $pair[1] ('=UXReportAssignmentDecision("'+$pair[0]+'ed")') 'ASSIGNMENT');$i++}
$f.Controls.Item('cmdAssignmentAccept').OnClick='=UXReportAssignmentDecision("Accepted")';$f.Controls.Item('cmdAssignmentReturn').OnClick='=UXReportAssignmentDecision("Returned")';$f.Controls.Item('cmdAssignmentCancel').OnClick='=UXReportAssignmentDecision("Cancelled")'
$f.Section(0).Height=10200;$app.DoCmd.Close(2,$f.Name,1)
# Shared conversation, live findings and calculations have their own embedded controls.
$f=ToolForm 'frmReportTools' 8500
foreach($name in @('txtReportID','txtDraftGuid','txtDraftVersion','txtDraftParent','txtDraftSaved','txtView')){$c=ToolControl $f 109 $name 0 0 100 100;$c.Visible=$false;if($name -in @('txtReportID','txtDraftVersion','txtDraftParent')){$c.DefaultValue='0'}}
[void](ToolControl $f 104 'uxToolFocus' 0 0 60 60 '' '')
[void](ToolControl $f 100 'lblToolStatus' 200 0 11800 300 '')
[void](ToolControl $f 100 'lblNoteIntro' 200 100 11800 600 'Report notes and replies - shared communication, independent of review or approval.' '' 'VIEW:Notes')
[void](ToolControl $f 100 'lblConversations' 200 850 3900 300 'Conversations - all contributors' '' 'VIEW:Notes')
$c=ToolControl $f 110 'lstMessages' 200 1250 3900 2600 '' '' 'VIEW:Notes';$c.ColumnCount=4;$c.ColumnHeads=$true;$c.ColumnWidths='0;1900;950;1000';$c.AfterUpdate='=UXMessageRead()'
[void](ToolControl $f 100 'lblConversation' 4300 850 7700 300 'Selected conversation and replies' '' 'VIEW:Notes')
$c=ToolControl $f 109 'txtConversation' 4300 1250 7700 2600 '' '' 'VIEW:Notes';$c.Locked=$true;$c.TabStop=$false;$c.BackColor=$app.Eval('UXColor("Neutral")')
[void](ToolControl $f 100 'lblDrafts' 200 4100 3900 300 'Your unfinished notes - not posted' '' 'VIEW:Notes')
$c=ToolControl $f 110 'lstDrafts' 200 4500 3900 1700 '' '' 'VIEW:Notes';$c.ColumnCount=3;$c.ColumnWidths='0;2400;1400'
[void](ToolControl $f 104 'cmdResume' 200 6350 3900 420 'Continue selected draft' '=UXMessageResume()' 'VIEW:Notes')
[void](ToolControl $f 104 'cmdNew' 4300 4100 3600 420 'Write a new note' '=UXMessageBegin(False)' 'VIEW:Notes')
[void](ToolControl $f 104 'cmdReply' 8100 4100 3900 420 'Reply to conversation' '=UXMessageBegin(True)' 'VIEW:Notes')
[void](ToolControl $f 100 'lblCompose' 4300 4700 7700 300 'Write a new report note' '' 'VIEW:Notes')
$c=ToolControl $f 109 'txtMessage' 4300 5100 7700 1600 '' '' 'VIEW:Notes';$c.EnterKeyBehavior=$true
[void](ToolControl $f 104 'cmdPost' 4300 6900 2350 450 'Post note / reply' '=UXMessagePost()' 'VIEW:Notes')
[void](ToolControl $f 104 'cmdKeepDraft' 6800 6900 2350 450 'Save draft' '=UXMessageKeep()' 'VIEW:Notes')
[void](ToolControl $f 104 'cmdDiscardDraft' 9300 6900 2700 450 'Discard draft' '=UXMessageDiscard()' 'VIEW:Notes')
[void](ToolControl $f 104 'cmdRefreshNotes' 200 6900 3900 450 'Refresh conversations' '=UXMessageRefresh()' 'VIEW:Notes')
[void](ToolControl $f 100 'lblNoteStatus' 200 7600 11800 700 '' '' 'VIEW:Notes')
[void](ToolControl $f 100 'lblValidation' 200 150 11800 550 '' '' 'VIEW:Validation')
$c=ToolControl $f 110 'lstFindings' 200 950 11800 2500 '' '' 'VIEW:Validation';$c.RowSourceType='Value List';$c.ColumnCount=4;$c.ColumnWidths='0;3000;1900;6750';$c.AfterUpdate='=UXToolsFinding()';$c.OnDblClick='=UXToolsFindingOpen()'
$c=ToolControl $f 109 'txtEvidence' 200 3700 11800 2900 '' '' 'VIEW:Validation';$c.Locked=$true
[void](ToolControl $f 104 'cmdFindingOpen' 200 6900 3800 450 'Go to affected report block' '=UXToolsFindingOpen()' 'VIEW:Validation')
[void](ToolControl $f 104 'cmdValidationRefresh' 4200 6900 3800 450 'Refresh checks' '=UXToolsValidation()' 'VIEW:Validation')
[void](ToolControl $f 104 'cmdValidationPopup' 8200 6900 3800 450 'Pop out validation issues' '=UXOpenValidation()' 'VIEW:Validation')
[void](ToolControl $f 100 'lblCalculations' 200 200 11800 500 'Trait calculations - pending editor values, saved values, and summary-group basis' '' 'VIEW:Calculations')
$c=ToolControl $f 109 'txtCalculations' 200 950 11800 5700 '' '' 'VIEW:Calculations';$c.Locked=$true
[void](ToolControl $f 104 'cmdCalculationRefresh' 200 6900 11800 450 'Refresh calculation details' '=UXToolsCalculations()' 'VIEW:Calculations')
$c=$app.CreateControl($f.Name,112,0,'','',0,0,12200,8200);$c.Name='uxReviewPanel';$c.SourceObject='Form.frmReportReviewPanel';$c.LinkMasterFields='';$c.LinkChildFields='';$c.Visible=$false;$c.BorderStyle=0
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
$c=$app.CreateControl($f.Name,112,0,'','',2850,4800,12600,7000);$c.Name='uxToolsSurface';$c.SourceObject='Form.frmReportTools';$c.LinkMasterFields='';$c.LinkChildFields='';$c.Visible=$false;$c.BorderStyle=0
$f.Controls.Item('uxPageNotes').Caption='Notes';$f.Controls.Item('uxPageNotes').OnClick='=UXToolsChoose("Notes")'
$f.Controls.Item('uxPageReview').Caption='Report tools';$f.Controls.Item('uxPageReview').OnClick='=UXSelectSection(8)'
foreach($v in @(@('cmdComments','Notes and replies','Notes'),@('cmdRoute','Review steps','Review'),@('cmdChecklist','Validation checks','Validation'),@('cmdCalculations','Trait calculations','Calculations'))){$f.Controls.Item($v[0]).Caption=$v[1];$f.Controls.Item($v[0]).OnClick='=UXToolsChoose("'+$v[2]+'")'}
$f.Controls.Item('uxReviewCommentGuide').Visible=$false;$f.Controls.Item('uxReviewCommentGuide').Tag=''
$app.DoCmd.Close(2,$f.Name,1)
# A modeless native popout can remain beside the report while its fields are edited.
$app.DoCmd.OpenForm('frmUXValidation',1);$f=$app.Forms.Item('frmUXValidation');$f.PopUp=$true;$f.Modal=$false;$f.BorderStyle=3;$f.AutoCenter=$true;$f.AutoResize=$false;$f.Controls.Item('cmdBack').Caption='Close validation window';$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmMyWork',1);$f=$app.Forms.Item('frmMyWork')
[void](ToolControl $f 104 'uxWorkNotesTab' 11000 1900 2300 450 'Report notes' '=UXMyWorkMode("Notes")')
[void](ToolControl $f 100 'lblReportNotes' 2850 2650 10450 550 'Conversations across this workspace - independent of review tasks' '' 'MWX:body')
$c=ToolControl $f 110 'lstReportNotes' 2850 3400 10450 3700 '' '' 'MWX:body';$c.ColumnCount=6;$c.ColumnHeads=$true;$c.ColumnWidths='0;0;3200;1700;3500;1800';$c.OnDblClick='=UXMyWorkOpenNote()'
[void](ToolControl $f 104 'cmdOpenReportNote' 2850 7500 5100 450 'Open selected report conversation' '=UXMyWorkOpenNote()' 'MWX:body')
[void](ToolControl $f 104 'cmdRefreshReportNotes' 8200 7500 5100 450 'Refresh report notes' '=UXMyWorkNotes()' 'MWX:body')
$app.DoCmd.Close(2,$f.Name,1)
foreach($name in @('frmReviewComments','frmUXReview','frmChecklist','frmUXNoteComposer')){if(@($app.CurrentProject.AllForms|ForEach-Object {$_.Name}) -contains $name){$app.DoCmd.DeleteObject(2,$name)}}
