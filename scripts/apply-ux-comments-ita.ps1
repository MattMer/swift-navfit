# Clear labels and reading/writing areas; existing storage, revision and attribution services remain authoritative.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
if(-not $script:swiftManagedBuild){foreach($m in @('modUXNotes','modReviewComments','modUXTraitGrid','modReportPages','modUXMyWork')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
function SetCommentRect($f,[string]$name,[int]$x,[int]$y,[int]$w,[int]$h){$c=$f.Controls.Item($name);foreach($v in @(@('Left',$x),@('Top',$y),@('Width',$w),@('Height',$h))){SetAccessProperty $c $v[0] $v[1]}}
function AddCommentLabel($f,[string]$name,[string]$caption,[int]$x,[int]$y,[int]$w,[int]$h){$c=$app.CreateControl($f.Name,100,0,'','',$x,$y,$w,$h);$c.Name=$name;$c.Caption=$caption;$c.FontName='Segoe UI';$c.FontSize=10;$c.Visible=$false;return $c}
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
