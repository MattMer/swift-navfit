$app.LoadFromText(5,'modUXDirectory',(Join-Path $root 'src/modUXDirectory.bas'))
NewScreen 'frmUXDirectory' 'Command member directory' 7600;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblScope' 400 950 12600 600 'Members are linked by SSN or deliberate confirmation. Name-only records remain provisional. Profile edits do not rewrite existing reports.')
[void](Control $f 109 'txtSearch' 400 1750 9500 450);$script:swiftControl.AfterUpdate='=UXDirectoryLoad()'
[void](Control $f 104 'cmdSearch' 10300 1750 2700 450 'Search members' '=UXDirectoryLoad()')
[void](Control $f 110 'lstMembers' 400 2400 12600 3200);$c=$script:swiftControl;$c.ColumnCount=5;$c.ColumnHeads=$true;$c.ColumnWidths='0;5600;1900;2400;2400';$c.OnDblClick='=UXDirectoryEdit(False)'
[void](Control $f 104 'cmdNew' 400 5900 4000 500 'Add member' '=UXDirectoryEdit(True)')
[void](Control $f 104 'cmdEdit' 4650 5900 4100 500 'Edit selected member' '=UXDirectoryEdit(False)')
[void](Control $f 104 'cmdClose' 9000 5900 4000 500 'Back to Resources' '=UXCloseResource("frmUXDirectory")')
$f.OnLoad='=UXDirectoryLoad()';SaveScreen $f 'frmUXDirectory'
NewScreen 'frmUXMemberEditor' 'Command member profile' 6900;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.ScrollBars=2
$items=@(@('txtName','Member display name (120 characters)'),@('txtReportName','Report name - Last, First MI Suffix (27 characters)'),@('txtRank','Rank/rate (5 characters)'),@('txtSSN','SSN - nine digits, or leave unentered'))
$y=1000;foreach($item in $items){[void](Control $f 100 ('lbl'+$item[0]) 400 $y 9600 350 $item[1]);[void](Control $f 109 $item[0] 400 ($y+400) 9600 450);$y+=950}
[void](Control $f 100 'lblStatus' 400 4950 9600 650 'Display name and report name are separately maintained. No name or SSN is truncated.')
[void](Control $f 104 'cmdSave' 400 5850 4650 500 'Save member profile' '=UXDirectorySave()')
[void](Control $f 104 'cmdCancel' 5350 5850 4650 500 'Cancel' '=UXDirectoryClose("frmUXMemberEditor")')
SaveScreen $f 'frmUXMemberEditor'
$app.DoCmd.OpenForm('frmUXResources',1);$f=$app.Forms.Item('frmUXResources')
$bottom=([int](($f.Controls | Where-Object {$_.ControlType -eq 104 -and $_.Visible -and $_.Name -notmatch '^uxNav'} | ForEach-Object {$_.Top+$_.Height} | Measure-Object -Maximum).Maximum))+450
[void](Control $f 104 'uxDirectory' 2850 $bottom 10450 550 'Command member directory' '=OpenScreen("frmUXDirectory")')
$f.Section(0).Height=[int]([Math]::Max($f.Section(0).Height,$bottom+1100));$f.ScrollBars=2;SaveScreen $f 'frmUXResources'
