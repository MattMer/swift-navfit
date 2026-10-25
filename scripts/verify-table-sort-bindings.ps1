param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/sort-bindings-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$retired=@((Get-Content (Join-Path $root 'scripts/build-manifest.json') -Raw|ConvertFrom-Json).retiredForms.PSObject.Properties.Name)
$app=New-Object -ComObject Access.Application;$count=0;$missing=@()
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 for($wait=0;$wait -lt 30 -and $app.Forms.Count -eq 0;$wait++){Start-Sleep -Milliseconds 100}
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 #Rule: every list showing two or more columns is sortable, by native header
 #clicks when it shows column heads and by clickable Sort headings otherwise.
 foreach($screen in @($app.CurrentProject.AllForms|ForEach-Object {$_.Name}|Where-Object {$retired -notcontains $_})){
  $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
  foreach($c in @($f.Controls|Where-Object {$_.ControlType -eq 110 -and $_.ColumnCount -gt 1})){
   $widths=@($c.ColumnWidths -split ';');$visible=0
   for($i=0;$i -lt $c.ColumnCount;$i++){if($i -ge $widths.Count -or $widths[$i] -eq '' -or [double]$widths[$i] -gt 0){$visible++}}
   if($visible -lt 2){continue}
   if($c.ColumnHeads){
    $text=if($f.HasModule){$f.Module.Lines(1,[Math]::Max(1,$f.Module.CountOfLines))}else{''}
    if($c.OnMouseDown -ne '[Event Procedure]' -or $text -notmatch ($c.Name+'_MouseDown') -or $text -notmatch ('UXSortNativeHeader\(Me.Name, "'+$c.Name+'"')){$missing+=($screen+'.'+$c.Name+' (native header)');continue}
   }else{
    $names=@($f.Controls|ForEach-Object {$_.Name})
    if(@(0..($c.ColumnCount-1)|Where-Object {$names -notcontains ('sort_'+$c.Name+'_'+$_) -or $f.Controls('sort_'+$c.Name+'_'+$_).OnClick -ne ('=UXSortList("'+$screen+'","'+$c.Name+'",'+$_+')')}).Count){$missing+=($screen+'.'+$c.Name+' (Sort headings)');continue}
   }
   $count++
  }
  $app.DoCmd.Close(2,$screen,2)
 }
 #Known gaps: lists allowed to lack sorting for now, each with a reason. Remove
 #an entry when its list gains sorting.
 $knownGaps=@()
 $unexpected=@($missing|Where-Object {$knownGaps -notcontains $_});$fixed=@($knownGaps|Where-Object {$missing -notcontains $_})
 if($unexpected.Count){throw ('Lists without sorting: '+($unexpected -join ', '))}
 if($fixed.Count){throw ('Known sorting gap is fixed; remove it from the list: '+($fixed -join ', '))}
 #Editable trait grid: each header sorts by the nearest visible bound field below it.
 $app.DoCmd.OpenForm('frmUXComparison',1);$grid=([string]$app.Forms.Item('frmUXComparison').Controls('uxTraitGrid').SourceObject) -replace '^Form\.','';$app.DoCmd.Close(2,'frmUXComparison',2)
 $app.DoCmd.OpenForm($grid,1);$f=$app.Forms.Item($grid);$headers=0
 foreach($h in @($f.Controls|Where-Object {$_.ControlType -eq 100 -and $_.Section -eq 1})){
  $best=$null;$distance=999999
  foreach($c in $f.Controls){if($c.ControlType -in @(109,111) -and $c.Section -eq 0 -and $c.ControlSource -and $c.Visible){$d=[math]::Abs($c.Left-$h.Left);if($d -lt $distance){$distance=$d;$best=$c}}}
  if($null -ne $best){if($h.OnClick -ne ('=UXSortGrid("'+$best.ControlSource+'")')){throw ('Missing grid heading: '+$h.Name)};$headers++;$count++}
 }
 if($headers -eq 0){throw 'Editable grid has no sortable headings'}
 $app.DoCmd.Close(2,$grid,2)
 $app.DoCmd.OpenForm('frmUXRSCATable',1);$f=$app.Forms.Item('frmUXRSCATable')
 foreach($key in @('grade','avg','count','date')){if($f.Controls('head'+$key).OnClick -ne ('=RSTableSort("'+$key+'")')){throw ('Missing RSCA heading: '+$key)};$count++}
 $app.DoCmd.Close(2,'frmUXRSCATable',2)
 $app.DoCmd.RunCommand(126)
 Write-Output ('PASS: '+$count+' sortable lists and headings verified against the sort rule')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
