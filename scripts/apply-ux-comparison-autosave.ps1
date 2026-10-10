# Finish bound-control events before the report transaction runs.
$app.LoadFromText(5,'modUXTraitGrid',(Join-Path $root 'src/modUXTraitGrid.bas'))
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
foreach($key in @('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39','PromotionRecommendation','AssignmentRecommendationA','AssignmentRecommendationB')){
    $f.Controls.Item($key).AfterUpdate='=UXGridQueue("'+$key+'")'
}
$app.DoCmd.Close(2,'frmUXTraitRows',1)
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison')
$f.OnTimer='=UXGridDrain()';$f.TimerInterval=0
$app.DoCmd.Close(2,'frmUXComparison',1)
