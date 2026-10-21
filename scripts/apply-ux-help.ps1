if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXHelp',(Join-Path $root 'src/modUXHelp.bas'))}
NewScreen 'frmUXHelp' 'Swift NAVFIT workspace guide' 8100;$f=$script:swiftForm;$f.PopUp=$true;$f.Width=13400
[void](Control $f 100 'lblTopics' 400 1100 3500 300 'Choose a workflow')
[void](Control $f 110 'lstTopics' 400 1550 3500 5500);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;3300';$c.AfterUpdate='=UXHelpSelected()';$c.TabIndex=0
[void](Control $f 100 'lblGuide' 4200 1100 8700 300 'How to use this workspace')
[void](Control $f 109 'txtGuide' 4200 1550 8700 5500);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2;$c.TabIndex=1
[void](Control $f 104 'cmdClose' 4200 7400 3000 450 'Close guide' '=UXCloseScreen("frmUXHelp")');$script:swiftControl.TabIndex=2
$f.OnLoad='=UXLoadHelp()';SaveScreen $f 'frmUXHelp'
$app.DoCmd.OpenForm('frmUXSettings',1);$f=$app.Forms.Item('frmUXSettings')
[void](Control $f 104 'cmdWorkspaceGuide' 2866 6700 9954 550 'Workspace guide and quick start' '=OpenScreen("frmUXHelp")')
$f.Section(0).Height=7500;SaveScreen $f 'frmUXSettings'
