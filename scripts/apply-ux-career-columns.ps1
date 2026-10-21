# Add existing career fields to the editable comparison's per-session cache UI.
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXTraitGrid',(Join-Path $root 'src/modUXTraitGrid.bas'))}
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
foreach($item in @(@('AssignmentRecommendationA',10450,'Career recommendation 1'),@('AssignmentRecommendationB',11850,'Career recommendation 2'))){
$c=$app.CreateControl($f.Name,100,1,'','',[int]$item[1],20,1350,900);$c.Name='heading'+$item[0];$c.Caption=$item[2];$c.FontName='Segoe UI';$c.FontSize=8
[void](Control $f 109 $item[0] $item[1] 40 1350 360);$c=$script:swiftControl;$c.ControlSource=$item[0];$c.AfterUpdate='=UXGridQueue("'+$item[0]+'")'
}
foreach($c in $f.Controls){$c.Tag=[string]$c.Left+':'+[string]$c.Width}
$f.Width=13250
SaveScreen $f 'frmUXTraitRows'
