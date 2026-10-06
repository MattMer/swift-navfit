$app.LoadFromText(5,'modUXDirectory',(Join-Path $root 'src/modUXDirectory.bas'))
$app.DoCmd.OpenForm('frmUXConfirmIdentity',1);$f=$app.Forms.Item('frmUXConfirmIdentity')
[void](Control $f 100 'lblMember' 400 3650 12600 350 'Directory member - choose the correct person')
[void](Control $f 111 'cboMember' 400 4050 12600 450);$c=$script:swiftControl
$c.ColumnCount=3;$c.ColumnWidths='0;9000;3000';$c.BoundColumn=1;$c.LimitToList=$true;$c.AfterUpdate='=UXHistoryIdentityPreview()'
$f.Controls.Item('lblReason').Top=4750;$f.Controls.Item('txtReason').Top=5150
$f.Controls.Item('lblStatus').Top=5900;$f.Controls.Item('cmdApply').Top=6600;$f.Controls.Item('cmdCancel').Top=6600
$f.Section(0).Height=7400;SaveScreen $f 'frmUXConfirmIdentity'
