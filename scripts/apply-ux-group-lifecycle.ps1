$app.LoadFromText(5,'modUXGroupLifecycle',(Join-Path $root 'src/modUXGroupLifecycle.bas'))
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups')
[void](Control $f 100 'uxFilterLabel' 2850 1950 2200 300 'Archive status')
[void](Control $f 111 'uxArchiveFilter' 5250 1900 3300 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='Active;Archived;All';$c.DefaultValue='"Active"';$c.LimitToList=$true;$c.AfterUpdate='=UXGroupList()';$c.ControlTipText='Filter active, archived or all summary groups'
$f.Controls('lstGroups').Top=2550;$f.Controls('lstGroups').Height=3300;$f.Controls('lstGroups').ColumnCount=6;$f.Controls('lstGroups').ColumnWidths='0;4400;1500;1000;1100;1700';$f.Controls('lstGroups').OnClick='=UXGroupSelection()';$f.Controls('lstGroups').AfterUpdate='=UXGroupSelection()'
$f.Controls('cmdOpen').Left=2850;$f.Controls('cmdOpen').Top=6100;$f.Controls('cmdOpen').Width=5000
$f.Controls('cmdNew').Left=8250;$f.Controls('cmdNew').Top=6100;$f.Controls('cmdNew').Width=5050
$i=0;foreach($action in @('Rename','Archive','Recover')){
[void](Control $f 104 ('ux'+$action) (2850+$i*3500) 6800 3200 450 ($action+' selected group') ('=UXGroupAction("'+$action+'")'));$i++
}
[void](Control $f 100 'uxLifecycleStatus' 2850 7450 10450 550 'Choose a group.')
$f.Section(0).Height=8300
SaveScreen $f 'frmUXGroups'
NewScreen 'frmUXRenameGroup' 'Rename summary group' 3700;$f=$script:swiftForm;$f.PopUp=$true;$f.Modal=$true;$f.Width=10000
$f.Controls('lblHeading').Width=9200
foreach($name in @('groupID','baseline')){[void](Control $f 109 $name 0 0 100 100);$script:swiftControl.Visible=$false}
[void](Control $f 100 'lblName' 400 1100 9000 300 'Group name')
[void](Control $f 109 'txtName' 400 1450 9000 450);$script:swiftControl.TabIndex=0;$script:swiftControl.ControlTipText='Summary group name, maximum 120 characters'
[void](Control $f 104 'cmdSave' 5200 2150 4200 500 'Save name' '=UXSaveGroupName()');$script:swiftControl.Default=$true
[void](Control $f 104 'cmdCancel' 400 2150 4200 500 'Cancel' '=UXCloseScreen("frmUXRenameGroup")');$script:swiftControl.Cancel=$true
[void](Control $f 100 'lblStatus' 400 2800 9000 550 'Renaming keeps reports, members, defaults and history.')
$f.OnLoad='=UXLoadRenameGroup()'
SaveScreen $f 'frmUXRenameGroup'
