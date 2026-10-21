# All group pages share a safe exit and explicit return to the group directory.
if(-not $script:swiftManagedBuild){foreach($m in @('modUXGroupNavigation','modUX')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
foreach($name in @('frmUXComparison','frmUXBulk','frmMajorWorkbench')){
 $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
 if($name -eq 'frmMajorWorkbench'){$f.Controls.Item('uxGroupMeta').Width=6800}
 if($name -eq 'frmUXBulk'){$f.Controls.Item('uxGroupTitle').Left=2850;$f.Controls.Item('uxGroupTitle').Width=6800}
 $c=$app.CreateControl($name,104,0,'','',10000,1550,3300,400)
 $c.Name='uxBackToGroups';$c.Caption='Back to summary groups';$c.OnClick='=UXBackToGroups()'
 $c.FontName='Segoe UI';$c.FontSize=10;$c.TabStop=$true
 try{$f.Controls.Item('cmdBack').Caption='Back to summary groups';$f.Controls.Item('cmdBack').OnClick='=UXBackToGroups()'}catch{}
 [void]$app.Eval('UXApplyVisual(Forms("'+$name+'"))')
 $app.DoCmd.Close(2,$name,1)
}
