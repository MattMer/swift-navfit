param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/property-bindings-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application;$checks=0
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 #Forms are source (src/forms), so these check the built forms directly.
 $app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary')
 if(@($f.Controls|ForEach-Object {$_.Name}) -contains 'cmdBulkFolders'){throw 'Retired cmdBulkFolders is still installed'};$checks++
 $app.DoCmd.Close(2,'frmLibrary',2)
 $app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison')
 foreach($name in @('cmdRetryCells','cmdDiscardCells')){
  $c=$f.Controls.Item($name)
  if($c.Top+$c.Height -gt $f.Section(0).Height){throw ('Recovery action outside section: '+$name)};$checks++
 }
 if($f.OnTimer -ne '=UXGridDrain()' -or $f.Module.Lines(1,$f.Module.CountOfLines) -notmatch 'If Not UXGridUnload\(Me\) Then Cancel = True'){throw 'Pending-cell events missing'};$checks++
 $app.DoCmd.Close(2,'frmUXComparison',2)
 $app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
 foreach($key in @('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39','PromotionRecommendation','AssignmentRecommendationA','AssignmentRecommendationB')){
  if($f.Controls.Item($key).AfterUpdate -ne ('=UXGridQueue("'+$key+'")')){throw ('Queue binding missing: '+$key)};$checks++
 }
 $app.DoCmd.Close(2,'frmUXTraitRows',2)
 Write-Output ('PASS: '+$checks+' retired-control, recovery-action bounds and native queue/unload bindings')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
