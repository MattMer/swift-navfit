# Additive schema 27. This stage can upgrade an existing verified schema-26 blank UI.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
$db=$app.CurrentDb();$names=@($db.TableDefs|ForEach-Object {$_.Name})
$mergeTables=@('DirectoryMerge','DirectoryMergeReport','DirectoryMergePrior')
$present=@($mergeTables|Where-Object {$names -contains $_})
if($present.Count -gt 0 -and $present.Count -ne 3){throw 'Incomplete directory-merge schema; no automatic repair'}
if($present.Count -eq 0){
 $schema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw|ConvertFrom-Json
 $w=$app.DBEngine.Workspaces.Item(0);$w.BeginTrans()
 try{foreach($sql in $schema.SQL|Where-Object {$_ -match 'DirectoryMerge'}){$db.Execute($sql,128)};$w.CommitTrans()}catch{$w.Rollback();throw}
}
$db.Execute('UPDATE AppMetadata SET SchemaVersion=27',128);$db.Close()
$app.DoCmd.OpenForm('frmUXDirectory',1);$f=$app.Forms.Item('frmUXDirectory')
if(@($f.Controls|ForEach-Object {$_.Name}) -contains 'cmdMerge'){$app.DeleteControl($f.Name,'cmdMerge')}
$edge=[int](($f.Controls|ForEach-Object {$_.Top+$_.Height}|Measure-Object -Maximum).Maximum)+300
$button=$app.CreateControl($f.Name,104,0,'','',2850,$edge,10450,500);$button.Name='cmdMerge';$button.Caption='Review possible duplicate Sailors';$button.OnClick='=UXDirectoryMergeOpen()';$button.BackColor=5948108;$button.ForeColor=16777215
SetAccessProperty ($f.Section(0)) 'Height' ($edge+1000);$app.DoCmd.Close(2,$f.Name,1)
$f=$app.CreateForm();$old=$f.Name;$f.Caption='Review possible duplicate Sailors';$f.Width=14500;$f.PopUp=$true;$f.Modal=$true;$f.AutoCenter=$true;$f.ScrollBars=2;$f.RecordSelectors=$false;$f.NavigationButtons=$false;$f.DividingLines=$false
function AddMergeControl($type,$name,$x,$y,$width,$height,$caption='',$event=''){
 $c=$app.CreateControl($f.Name,$type,0,'','',$x,$y,$width,$height);$c.Name=$name;$c.FontName='Segoe UI';$c.FontSize=10
 if($type -eq 100 -or $type -eq 104){$c.Caption=$caption}
 if($event){$c.OnClick=$event}
 return $c
}
$c=AddMergeControl 104 'cmdWindowHelp' 13700 150 500 400 '?' '=UXWindowHelp("frmUXDirectoryMerge")';$c.ControlTipText='Help for this window';$c.FontBold=$true
$f.OnResize='=UXHelpResize("frmUXDirectoryMerge","")'
[void](AddMergeControl 100 'lblInfo' 300 200 13000 750 'Review two profiles for the same Sailor. Surname matches are suggestions only. Conflicting identifiers block merging. The original profile stays archived; saved report content and history are preserved.')
[void](AddMergeControl 100 'lblSource' 300 1000 6700 350 'Original member - archive after moving links')
[void](AddMergeControl 100 'lblTarget' 7500 1000 6700 350 'Surviving member - keep this profile')
foreach($pair in @(@('cboSource',300),@('cboTarget',7500))){$c=AddMergeControl 111 $pair[0] $pair[1] 1400 6700 430;$c.RowSourceType='Table/Query';$c.ColumnCount=2;$c.ColumnWidths='0;6400';$c.BoundColumn=1;$c.LimitToList=$true;$c.AfterUpdate='=UXDirectoryMergeChanged()'}
[void](AddMergeControl 100 'lblCandidates' 300 1900 13900 350 'Possible surname matches - double-click to choose a survivor, or use the survivor list for changed names')
$c=AddMergeControl 110 'lstCandidates' 300 2300 13900 1100;$c.ColumnCount=3;$c.ColumnHeads=$true;$c.ColumnWidths='0;10500;2500';$c.OnDblClick='=UXDirectoryMergeCandidate()'
[void](AddMergeControl 104 'cmdPreview' 300 3500 13900 450 'Preview identities and all affected links' '=UXDirectoryMergePreview()')
$c=AddMergeControl 109 'txtPreview' 300 4050 13900 3400;$c.Locked=$true;$c.TabStop=$true;$c.ScrollBars=2;$c.EnterKeyBehavior=$true
[void](AddMergeControl 100 'lblReason' 300 7550 13900 350 'Why are these profiles the same Sailor? Required; do not include SSNs or DoD IDs')
[void](AddMergeControl 109 'txtReason' 300 7950 13900 500)
$c=AddMergeControl 104 'cmdApply' 300 8550 13900 500 'Apply reviewed merge' '=UXDirectoryMergeApply()';$c.Enabled=$false;$c.BackColor=5948108;$c.ForeColor=16777215
[void](AddMergeControl 100 'lblHistory' 300 9200 13900 350 'Merge history - undo restores original links only if no later changes affect the members')
$c=AddMergeControl 110 'lstHistory' 300 9600 13900 1300;$c.ColumnCount=5;$c.ColumnHeads=$true;$c.ColumnWidths='1000;2400;2400;2400;4200'
[void](AddMergeControl 104 'cmdUndo' 300 11000 6700 450 'Undo selected merge' '=UXDirectoryMergeUndo()')
[void](AddMergeControl 104 'cmdClose' 7500 11000 6700 450 'Back to member directory' '=UXDirectoryClose("frmUXDirectoryMerge")')
SetAccessProperty ($f.Section(0)) 'Height' 11800
$f.OnLoad='=UXDirectoryMergeLoad()';$app.DoCmd.Save(2,$old);$app.DoCmd.Close(2,$old,1);$app.DoCmd.Rename('frmUXDirectoryMerge',2,$old)
