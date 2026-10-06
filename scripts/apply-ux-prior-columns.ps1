$app.LoadFromText(5,'modUXTraitGrid',(Join-Path $root 'src/modUXTraitGrid.bas'))
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
foreach($item in @(@('PriorITA',13950,1000,'Prior ITA'),@('PriorPromotion',14950,2100,'Prior promotion'))){
$c=$app.CreateControl($f.Name,100,1,'','',[int]$item[1],20,[int]$item[2],900);$c.Name='heading'+$item[0];$c.Caption=$item[3];$c.FontName='Segoe UI';$c.FontSize=8
[void](Control $f 109 $item[0] $item[1] 40 $item[2] 360);$c=$script:swiftControl;$c.ControlSource=$item[0];$c.Locked=$true;$c.TabStop=$false
}
foreach($c in $f.Controls){$c.Tag=[string]$c.Left+':'+[string]$c.Width};$f.Width=17050
SaveScreen $f 'frmUXTraitRows'
