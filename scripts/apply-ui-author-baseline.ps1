# Author-reviewed UI changes and dependency-audited presentation cleanup.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
$baseline=Get-Content (Join-Path $root 'src/ui-author-baseline.json') -Raw|ConvertFrom-Json
$names=@(@($baseline.properties|ForEach-Object {$_.form})+@($baseline.add_controls|ForEach-Object {$_.form})+@($baseline.remove_controls|ForEach-Object {$_.form})+'frmLibrary'|Sort-Object -Unique)
foreach($formName in $names){
 $app.DoCmd.OpenForm($formName,1);$f=$app.Forms.Item($formName)
 if($formName -eq 'frmLibrary' -and -not @($f.Controls|ForEach-Object {$_.Name}).Contains('cboDestination')){
  [void](Control $f 111 'cboDestination' 10 10 100 100);$c=$script:swiftControl;$c.Visible=$false;$c.ColumnCount=2;$c.ColumnWidths='0;2500';$c.RowSourceType='Table/Query';$c.LimitToList=$true
 }
 foreach($entry in @($baseline.add_controls|Where-Object {$_.form -eq $formName})){
  if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains($entry.control)){[void](Control $f $entry.type $entry.control $entry.properties.Left $entry.properties.Top $entry.properties.Width $entry.properties.Height $entry.properties.Caption)}
  $c=$f.Controls.Item($entry.control)
  foreach($p in $entry.properties.PSObject.Properties){try{SetAccessProperty $c $p.Name $p.Value}catch{throw ($formName+'.'+$entry.control+'.'+$p.Name+': '+$_.Exception.Message)}}
 }
 foreach($entry in @($baseline.properties|Where-Object {$_.form -eq $formName})){
  $target=if($entry.control){$f.Controls.Item($entry.control)}else{$f}
  try{SetAccessProperty $target $entry.property $entry.after}catch{throw ($formName+'.'+$entry.control+'.'+$entry.property+': '+$_.Exception.Message)}
 }
 foreach($entry in @($baseline.remove_controls|Where-Object {$_.form -eq $formName})){
  if(@($f.Controls|ForEach-Object {$_.Name}).Contains($entry.control)){$app.DeleteControl($formName,$entry.control)}
 }
 SaveScreen $f $formName
}


Write-Output ('Author UI baseline applied: '+$baseline.properties.Count+' property changes.')
