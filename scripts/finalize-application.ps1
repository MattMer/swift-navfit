# Final form ownership: route cleanup runs after presentation/event decorators.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
# Canonical history shares the existing operations/undo destination.
$app.DoCmd.OpenForm('frmOperations',1);$f=$app.Forms.Item('frmOperations')
$bottom=[int](($f.Controls|ForEach-Object {$_.Top+$_.Height}|Measure-Object -Maximum).Maximum)+300
$c=$app.CreateControl($f.Name,100,0,'','',400,$bottom,12600,350);$c.Name='lblAudit';$c.Caption='Report activity - domain evidence and operation correlation'
$c=$app.CreateControl($f.Name,110,0,'','',400,($bottom+500),12600,2500);$c.Name='lstAudit';$c.ColumnCount=6;$c.ColumnWidths='0;2000;2700;2200;900;4400';$c.ColumnHeads=$true;$c.AfterUpdate='=AuditDetails()'
$f.Section(0).Height=$bottom+3300;$f.ScrollBars=2
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary')
foreach($name in @('cmdAddFolder','cmdMoveFolder','cmdFolderHelp','cboDestination')){
 if(@($f.Controls|ForEach-Object {$_.Name}) -contains $name){$app.DeleteControl($f.Name,$name)}
}
$f.Controls.Item('cmdBulkFolders').Caption='Folder tools'
$f.Controls.Item('cmdBulkFolders').ControlTipText='Organizational folder moves; summary-group membership and official report content stay separate.'
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmBulkFolders',1);$f=$app.Forms.Item('frmBulkFolders')
$f.Caption='Organizational folder tools'
$f.Controls.Item('lblScope').Caption='Select active drafts and preview their organizational folder moves. These moves do not change summary-group membership or saved report content.'
$f.Controls.Item('cmdLibrary').Caption='Back to reports'
$app.DoCmd.Close(2,$f.Name,1)
$formEvents=@('OnOpen','OnLoad','OnClose','OnActivate','OnResize','OnTimer','OnUnload')
$controlEvents=@('OnClick','OnDblClick','AfterUpdate','BeforeUpdate','OnEnter','OnExit','OnChange','OnMouseDown')
$retired=@($swiftBuildPlan.Manifest.retiredForms.PSObject.Properties.Name)
foreach($name in @($app.CurrentProject.AllForms|ForEach-Object {$_.Name})){
 if($retired -contains $name){continue}
 $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
 foreach($obj in @(,$f)+@($f.Controls)){
  $properties=if($obj.Name -eq $f.Name){$formEvents}else{$controlEvents}
  foreach($property in $properties){
   try{$old=[string]$obj.Properties.Item($property).Value}catch{continue}
   $new=$old.Replace('OpenScreen("frmHelp")','OpenScreen("frmUXHelp")').Replace('UXNavigate("frmHelp")','UXNavigate("frmUXHelp")')
   if($new -ne $old){$obj.Properties.Item($property).Value=$new}
   foreach($target in $retired){if($new -match ('["!]' + [regex]::Escape($target) + '(?:["!]|\b)')){throw ('Unresolved retired route '+$target+' in '+$name+'.'+$obj.Name+'.'+$property)}}
  }
  if($obj.Name -ne $f.Name -and $obj.ControlType -eq 112){
   $child=([string]$obj.SourceObject) -replace '^Form\.',''
   if($retired -contains $child){throw ('Retired subform still used: '+$name+'.'+$obj.Name)}
  }
 }
 $app.DoCmd.Close(2,$name,1)
}
foreach($name in $retired){
 if(@($app.CurrentProject.AllForms|ForEach-Object {$_.Name}) -contains $name){$app.DoCmd.DeleteObject(2,$name)}
}
