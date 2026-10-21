# Discover populated native folders even when report rank/rate needs completion.
if(-not $script:swiftManagedBuild){foreach($m in @('modImportReview','modUXNativeGroups')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmUXNativeGroups',1);$f=$app.Forms.Item('frmUXNativeGroups')
$f.Controls('lstGroups').Height=2050
$f.Controls('lblCentury').Caption='Counseling date century (required only for ambiguous source dates)'
foreach($name in @('lblImportGrade','cboImportGrade','cmdAssignImportGrade')){try{$app.DeleteControl($f.Name,$name)}catch{}}
$c=$app.CreateControl($f.Name,100,0,'','',400,4300,4700,400);$c.Name='lblImportGrade';$c.Caption='Paygrade for unresolved source rank/rate';$c.FontName='Segoe UI';$c.FontSize=10
$c=$app.CreateControl($f.Name,111,0,'','',5200,4300,2000,400);$c.Name='cboImportGrade';$c.RowSourceType='Value List';$c.RowSource='E1;E2;E3;E4;E5;E6;E7;E8;E9;W1;W2;W3;W4;W5;O1;O2;O3;O4;O5;O6;O7;O8;O9;O10';$c.LimitToList=$true;$c.FontName='Segoe UI';$c.FontSize=10
$c=$app.CreateControl($f.Name,104,0,'','',7500,4300,5500,400);$c.Name='cmdAssignImportGrade';$c.Caption='Assign paygrade to selected groups';$c.OnClick='=UXAssignNativeGroupGrade()';$c.FontName='Segoe UI';$c.FontSize=10
$c.BackColor=$f.Controls('cmdPreview').BackColor;$c.ForeColor=$f.Controls('cmdPreview').ForeColor
$app.DoCmd.Close(2,$f.Name,1)
