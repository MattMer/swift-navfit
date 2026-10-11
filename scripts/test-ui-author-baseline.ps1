param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$AllowGroupDirectoryLayout)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/author-ui-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$spec=Get-Content (Join-Path $root 'src/ui-author-baseline.json') -Raw|ConvertFrom-Json
$app=New-Object -ComObject Access.Application;$checks=0
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 if(-not $app.CurrentProject.AllForms.Item('frmHome').IsLoaded){throw 'Startup failed'}
 foreach($n in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$n,2)}
 foreach($name in @(@($spec.properties|ForEach-Object {$_.form})+@($spec.remove_controls|ForEach-Object {$_.form})|Sort-Object -Unique)){
  $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
  foreach($entry in @($spec.properties|Where-Object {$_.form -eq $name -and $_.property -ne 'TabIndex'})){
   if($AllowGroupDirectoryLayout -and $name -eq 'frmMajorWorkbench' -and $entry.control -eq 'lblStatus' -and $entry.property -in @('Top','Height')){continue}
   $c=if($entry.control){$f.Controls.Item($entry.control)}else{$f};$actual=$c.($entry.property)
   if($actual -ne $entry.after){throw ($name+'.'+$entry.control+'.'+$entry.property+' differs: '+$actual+' expected '+$entry.after)};$checks++
  }
  foreach($entry in @($spec.remove_controls|Where-Object {$_.form -eq $name})){
   if(@($f.Controls|ForEach-Object {$_.Name}).Contains($entry.control)){throw ('Obsolete control remains: '+$name+'.'+$entry.control)};$checks++
  }
  $app.DoCmd.Close(2,$name,2)
 }
 foreach($name in @('frmLibrary','frmUXGroups','frmMyWork','frmUXResources','frmUXImportChoose','frmUXMemberHistory')){$app.DoCmd.OpenForm($name);$checks++}
 $app.Forms.Item('frmUXGroups').Controls.Item('lstGroups').Value=$app.Eval('DMin("GroupID","SummaryGroup")')
 if(-not $app.Eval('UXOpenGroup()')){throw 'Group failed to open'};$checks++
 if(-not $app.CurrentProject.AllForms.Item('frmUXComparison').IsLoaded){throw 'Group did not open directly to editable members'};$checks++
 if($app.CurrentProject.AllForms.Item('frmMajorWorkbench').IsLoaded -and $app.Forms.Item('frmMajorWorkbench').Visible){throw 'Legacy workbench was displayed on the direct member route'};$checks++
 $app.DoCmd.Close(2,'frmUXComparison',2)
 Write-Output ('PASS: '+$checks+' author properties, obsolete-control absence, startup and populated navigation checks')
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2)}

