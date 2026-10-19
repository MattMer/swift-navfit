param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/property-bindings-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application;$checks=0
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 $source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
 Invoke-Expression ([regex]::Match($source,'(?s)    function Control\(.*?(?=    NewScreen ''frmSetup'')').Value)
 . (Join-Path $PSScriptRoot 'apply-ui-author-baseline.ps1')
 . (Join-Path $PSScriptRoot 'apply-ui-author-baseline.ps1')
 $baseline=Get-Content (Join-Path $root 'src/ui-author-baseline.json') -Raw|ConvertFrom-Json
 foreach($formName in @($baseline.properties.form+$baseline.add_controls.form|Sort-Object -Unique)){
  $app.DoCmd.OpenForm($formName,1);$f=$app.Forms.Item($formName)
  foreach($entry in @($baseline.properties|Where-Object {$_.form -eq $formName})){
   $target=if($entry.control){$f.Controls.Item($entry.control)}else{$f}
   if($target.Properties.Item($entry.property).Value -ne $entry.after){throw ('Baseline differs: '+$formName+'.'+$entry.control+'.'+$entry.property)};$checks++
  }
  foreach($entry in @($baseline.add_controls|Where-Object {$_.form -eq $formName})){
   foreach($property in $entry.properties.PSObject.Properties){
    if($f.Controls.Item($entry.control).Properties.Item($property.Name).Value -ne $property.Value){throw ('Added control differs: '+$entry.control+'.'+$property.Name)};$checks++
   }
  }
  $app.DoCmd.Close(2,$formName,2)
 }
 . (Join-Path $PSScriptRoot 'apply-ux-comparison-autosave.ps1')
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
 Write-Output ('PASS: '+$checks+' author-layout values, repeat application, recovery-action bounds and native queue/unload bindings')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
