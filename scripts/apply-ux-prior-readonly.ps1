# Prior evaluation values are reference-only: no mouse focus or keyboard tab stop.
if(-not $script:swiftManagedBuild){foreach($m in @('modUXVisual','modUXTraitGrid','modUXFeedback')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
foreach($name in @('PriorITA','PriorPromotion')){
 $c=$f.Controls.Item($name);$c.Locked=$true;$c.Enabled=$false;$c.TabStop=$false
 $c.BackColor=$app.Eval('UXColor("Neutral")');$c.ForeColor=$app.Eval('UXColor("Muted")')
 $f.Controls.Item('heading'+$name).ControlTipText='Read-only prior evaluation reference. Current report edits use the trait and promotion columns to the left.'
}
$app.DoCmd.Close(2,'frmUXTraitRows',1)
