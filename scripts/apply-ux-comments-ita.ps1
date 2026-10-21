# Clear labels and reading/writing areas; existing storage, revision and attribution services remain authoritative.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
if(-not $script:swiftManagedBuild){foreach($m in @('modUXNotes','modReviewComments','modUXTraitGrid','modReportPages','modUXMyWork')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
function SetCommentRect($f,[string]$name,[int]$x,[int]$y,[int]$w,[int]$h){$c=$f.Controls.Item($name);foreach($v in @(@('Left',$x),@('Top',$y),@('Width',$w),@('Height',$h))){SetAccessProperty $c $v[0] $v[1]}}
function AddCommentLabel($f,[string]$name,[string]$caption,[int]$x,[int]$y,[int]$w,[int]$h){$c=$app.CreateControl($f.Name,100,0,'','',$x,$y,$w,$h);$c.Name=$name;$c.Caption=$caption;$c.FontName='Segoe UI';$c.FontSize=10;$c.Visible=$false;return $c}
$app.DoCmd.OpenForm('frmUXEditorFields',1);$f=$app.Forms.Item('frmUXEditorFields')
foreach($name in @('uxNotesIntro','uxNotesPosted','uxNotesRead','uxNotesDraftLabel')){[void](AddCommentLabel $f $name '' 200 100 8000 300)}
$c=$app.CreateControl($f.Name,104,0,'','',3800,800,5000,450);$c.Name='uxNotesComments';$c.Caption='Open comments and replies';$c.OnClick='=CommentsFromEditor()';$c.FontName='Segoe UI';$c.FontSize=10;$c.Visible=$false
$f.Controls.Item('uxNoteDrafts').AfterUpdate='=UXNoteDraftSelection()'
$f.Controls.Item('uxNoteDetails').BackColor=$app.Eval('UXColor("Neutral")')
$app.DoCmd.Close(2,'frmUXEditorFields',1)
$app.DoCmd.OpenForm('frmUXNoteComposer',1);$f=$app.Forms.Item('frmUXNoteComposer');$f.Caption='Write a report note';$f.Controls.Item('lblHeading').Caption='Write a report note';$f.Width=10400
[void](AddCommentLabel $f 'lblNoteText' 'Note text - visible to contributors after you post' 400 1450 9600 350);$f.Controls.Item('lblNoteText').Visible=$true
SetCommentRect $f 'lblContributor' 400 1000 9600 350
SetCommentRect $f 'txtNote' 400 1850 9600 2600
$f.Controls.Item('txtNote').EnterKeyBehavior=$true
$f.Controls.Item('cmdPost').Caption='Post note';$f.Controls.Item('cmdPost').FontBold=$true
$f.Controls.Item('cmdKeep').Caption='Save draft and close';$f.Controls.Item('cmdDiscard').Caption='Discard draft'
[void]$app.Eval('UXApplyVisual(Forms("frmUXNoteComposer"))');$app.DoCmd.Close(2,'frmUXNoteComposer',1)
$app.DoCmd.OpenForm('frmReviewComments',1);$f=$app.Forms.Item('frmReviewComments');$f.Caption='Report comments and replies'
$f.Controls.Item('lblHeading').Caption='Report comments and replies'
$f.Controls.Item('lblScope').Caption='Read feedback from all contributors, start a new comment, or reply to a conversation. Comments stay separate from the official narrative.'
SetCommentRect $f 'lblHeading' 2850 750 10450 500
SetCommentRect $f 'lblScope' 2850 1300 10450 600
SetCommentRect $f 'lblContext' 2850 1950 10450 550
[void](AddCommentLabel $f 'lblCommentList' '1. Choose a conversation - all contributors' 2850 2600 6800 350);$f.Controls.Item('lblCommentList').Visible=$true
 $c=$app.CreateControl($f.Name,104,0,'','',9850,2520,3450,400);$c.Name='cmdRefreshComments';$c.Caption='Refresh comments';$c.OnClick='=UXRefreshComments()'
SetCommentRect $f 'lstThreads' 2850 3000 10450 1600
$f.Controls.Item('lstThreads').ColumnHeads=$true;$f.Controls.Item('lstThreads').AfterUpdate='=ShowReviewThread()'
SetCommentRect $f 'lblHistory' 2850 4750 10450 350;$f.Controls.Item('lblHistory').Caption='2. Read the selected conversation and replies'
SetCommentRect $f 'txtHistory' 2850 5150 10450 2300;$f.Controls.Item('txtHistory').Locked=$true;$f.Controls.Item('txtHistory').BackColor=$app.Eval('UXColor("Neutral")');$f.Controls.Item('txtHistory').TabStop=$false
$c=$app.CreateControl($f.Name,104,0,'','',2850,7650,3300,450);$c.Name='cmdNewComment';$c.Caption='Write new comment';$c.OnClick='=UXBeginComment(False)'
SetCommentRect $f 'cmdReply' 6350 7650 3300 450;$f.Controls.Item('cmdReply').Caption='Reply to selected';$f.Controls.Item('cmdReply').OnClick='=UXBeginComment(True)'
SetCommentRect $f 'lblCompose' 2850 8250 10450 350
SetCommentRect $f 'txtCompose' 2850 8650 10450 1300;$f.Controls.Item('txtCompose').EnterKeyBehavior=$true
SetCommentRect $f 'cmdPost' 2850 10150 3300 450;$f.Controls.Item('cmdPost').OnClick='=UXSendComment()';$f.Controls.Item('cmdPost').FontBold=$true
SetCommentRect $f 'cmdResolve' 6350 10150 3300 450;$f.Controls.Item('cmdResolve').Caption='Mark conversation resolved'
SetCommentRect $f 'cmdReopen' 9850 10150 3450 450;$f.Controls.Item('cmdReopen').Caption='Reopen conversation'
[void](AddCommentLabel $f 'lblCommentStatus' '' 2850 10800 10450 600);$f.Controls.Item('lblCommentStatus').Visible=$true
SetCommentRect $f 'cmdEditor' 2850 11550 5000 450;$f.Controls.Item('cmdEditor').Caption='Back to report'
SetCommentRect $f 'cmdMyWork' 8350 11550 4950 450
$f.Section(0).Height=12250;$f.ScrollBars=2
[void]$app.Eval('UXApplyVisual(Forms("frmReviewComments"))');$app.DoCmd.Close(2,'frmReviewComments',1)
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
$c=AddCommentLabel $f 'uxReviewCommentGuide' 'Comments and replies: read feedback from all contributors, write a new comment, or reply to a selected conversation. Review steps manages the separate approval route. Use Report notes for working notes and unfinished note drafts.' 2850 4850 11800 900
$c.Tag='LOCAL:8'
$f.Controls.Item('uxPageNotes').Caption='Report notes'
$f.Controls.Item('cmdComments').Caption='Comments and replies';$f.Controls.Item('cmdComments').FontBold=$true
$f.Controls.Item('cmdRoute').Caption='Review steps';$f.Controls.Item('cmdChecklist').Caption='Validation checks';$f.Controls.Item('cmdCalculations').Caption='Trait calculations'
$app.DoCmd.Close(2,'frmEditor',1)
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows');$c=$f.Controls.Item('ITA')
$c.FormatConditions.Delete()
$condition=$c.FormatConditions.Add(1,[Type]::Missing,'UXITAChange([ITA],[PriorITA])=1');$condition.BackColor=$app.Eval('UXColor("Success")');$condition.ForeColor=$app.Eval('UXColor("SuccessInk")')
$condition=$c.FormatConditions.Add(1,[Type]::Missing,'UXITAChange([ITA],[PriorITA])=-1');$condition.BackColor=$app.Eval('UXColor("Warning")');$condition.ForeColor=$app.Eval('UXColor("WarningInk")')
$c.ControlTipText='Compared with this sailor''s prior ITA: green = higher; yellow = lower. Equal or no prior ITA keeps the normal appearance.'
$app.DoCmd.Close(2,'frmUXTraitRows',1)
