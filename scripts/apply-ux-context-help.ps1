# Consolidate legacy help routes and install contextual help on standalone windows.
foreach($m in @('modUXHelp','modWorkspace','modUXViewport','modUXGroupPages','modSwift','modUX')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$names=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
$helpCount=0
foreach($name in $names){
 if($name -in @('frmHelp','frmUXHelp','frmUXDecision')){continue}
 $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
 # Embedded field/table forms do not represent windows.
 if(-not $f.Caption -or $name -match 'Fields|Surface|Rows|Table|GridRows'){ $app.DoCmd.Close(2,$name,2);continue }
 foreach($c in @($f.Controls)){
  if($c.ControlType -eq 104 -and $c.OnClick -match 'ShowHelp|OpenScreen\("frmHelp"\)|OpenScreen\("frmUXHelp"\)'){
   $c.OnClick='=UXWindowHelp("'+$name+'")'
   if($c.Name -ne 'cmdWorkspaceGuide'){$c.Visible=$false}
  }
 }
 $exists=@($f.Controls|ForEach-Object {$_.Name}).Contains('cmdWindowHelp')
 if(-not $exists){$c=$app.CreateControl($name,104,0,'','',($f.Width-800),150,500,400);$c.Name='cmdWindowHelp'}
 $c=$f.Controls.Item('cmdWindowHelp');$c.Caption='?';$c.FontSize=11;$c.FontBold=$true;$c.ControlTipText='Help for this window';$c.OnClick='=UXWindowHelp("'+$name+'")'
 $old=[string]$f.OnResize
 if($old -eq '[Event Procedure]'){throw ('Preserve native Resize handler explicitly: '+$name)}
 if($old -notlike '=UXHelpResize*'){
  $old=$old.TrimStart('=').Replace('"','""')
  $f.OnResize='=UXHelpResize("'+$name+'","'+$old+'")'
 }
 $app.DoCmd.Close(2,$name,1);$helpCount++
}
Write-Output ('Context help installed on '+$helpCount+' windows; legacy help actions use the workspace guide.')
