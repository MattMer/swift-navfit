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
foreach($name in @('cmdAddFolder','cmdMoveFolder','cmdFolderHelp','cboDestination','cmdBulkFolders')){
 if(@($f.Controls|ForEach-Object {$_.Name}) -contains $name){$app.DeleteControl($f.Name,$name)}
}
$app.DoCmd.Close(2,$f.Name,1)
# Retain nonvisual compatibility controls for protected return services; the
# supported user flow is exclusively selected-group export.
$app.DoCmd.OpenForm('frmNAVFITExchange',1);$f=$app.Forms.Item('frmNAVFITExchange')
foreach($c in $f.Controls){$c.Visible=$false;$c.Left=0;$c.Top=0;$c.Width=500;$c.Height=350;if($c.ControlType -in @(104,106,109,110,111)){$c.TabStop=$false}}
$f.Caption='Export summary group to NAVFIT98A';$f.Width=13400;$f.Section(0).Height=5800;$f.ScrollBars=0;$f.PopUp=$false;$f.Modal=$false;$f.CloseButton=$true
foreach($spec in @(@('lblHeading',400,250,12600,600),@('lblScope',400,1000,12600,650),@('lblGroup',400,1750,12600,450),@('lbltxtExportPath',400,2350,2900,350),@('txtExportPath',3500,2300,7300,450),@('cmdtxtExportPath',11100,2300,1900,450),@('cmdExportGroup',400,3000,12600,550),@('txtStatus',400,3800,12600,1000),@('cmdBack',400,5100,12600,450),@('cmdWindowHelp',12500,250,500,450))){
 $c=$f.Controls.Item($spec[0]);$c.Visible=$true;$c.Left=$spec[1];$c.Top=$spec[2];$c.Width=$spec[3];$c.Height=$spec[4];if($c.ControlType -eq 104){$c.TabStop=$true}
}
$f.Controls('lblHeading').Caption='Export summary group to NAVFIT98A'
$f.Controls('lblScope').Caption='Export all saved reports in this summary group. Validation findings and report status do not block export. Pending edits are kept separately; this does not approve or submit reports.'
$f.Controls('lblGroup').Caption='Summary group';$f.Controls('lbltxtExportPath').Caption='Export directory'
$f.Controls('txtExportPath').Locked=$false;$f.Controls('txtExportPath').TabStop=$true;$f.Controls('txtExportPath').AfterUpdate='=UXPreviewNAVFITGroup()'
$f.Controls('cmdtxtExportPath').Caption='Choose folder';$f.Controls('cmdtxtExportPath').OnClick='=UXChooseNAVFITGroupFolder()'
$f.Controls('cmdExportGroup').Caption='Export summary group to NAVFIT98A'
$f.Controls('cmdBack').Caption='Close';$f.Controls('cmdBack').OnClick='=UXCloseScreen("frmNAVFITExchange")'
$f.OnUnload='=UXCloseNAVFITExport()'
$app.DoCmd.Close(2,$f.Name,1)
# End summary-group export surface.
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
