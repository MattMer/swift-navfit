param([string]$FixtureManifest='planning/SCALE_R102_FIXTURE.json')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$fixture=Get-Content (Join-Path $root $FixtureManifest) -Raw|ConvertFrom-Json
$scratch=Join-Path $root ('build/tests/scale-measure-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $fixture.Database -Destination $scratch
$app=New-Object -ComObject Access.Application;$results=@()
function Timed([string]$expression){$watch=[Diagnostics.Stopwatch]::StartNew();$value=$app.Eval($expression);$watch.Stop();if($value -eq $false){throw ('Failed: '+$expression)};[double]$watch.Elapsed.TotalMilliseconds}
function Summary($samples){$ordered=@($samples|Sort-Object);$middle=[int][Math]::Floor($ordered.Count/2);$median=if($ordered.Count%2){$ordered[$middle]}else{($ordered[$middle-1]+$ordered[$middle])/2};@{Count=$ordered.Count;MedianMs=[Math]::Round($median,2);MeanMs=[Math]::Round(($ordered|Measure-Object -Average).Average,2);MinMs=[Math]::Round($ordered[0],2);MaxMs=[Math]::Round($ordered[-1],2)}}
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 $app.LoadFromText(5,'ScaleHarness',(Join-Path $root 'tests/ScaleHarness.bas'));$app.DoCmd.RunCommand(126)
 $app.DoCmd.OpenForm('frmUXGroups');$startHistory=$app.Eval('ScaleHistoryCount()');$startRevisions=$app.Eval('DCount("*","ReportRevision")')
 foreach($g in $fixture.Groups){
  $opens=@();for($repeat=0;$repeat -lt 3;$repeat++){$opens+=Timed ('ScaleOpen('+$g.GroupID+')');if($app.Eval('ScaleLoadedRows()') -ne $g.Size){throw 'Loaded row count mismatch'};if($repeat -lt 2){if(-not $app.Eval('UXCloseComparison()')){throw 'Close failed'}}}
  $positions=@(0,[int][Math]::Floor(($g.Size-1)/4),[int][Math]::Floor(($g.Size-1)/2),[int][Math]::Floor(3*($g.Size-1)/4),[int]($g.Size-1))
  $moves=@();for($repeat=0;$repeat -lt 3;$repeat++){foreach($ordinal in $positions){$moves+=Timed ('ScaleNavigate('+$ordinal+')')}}
  $saves=@();for($i=0;$i -lt 10;$i++){$ordinal=[int][Math]::Floor($i*($g.Size-1)/9);$saves+=Timed ('ScaleSave('+$ordinal+')')}
  $editorOpens=@();$editorCloses=@();foreach($ordinal in @($positions[0],$positions[2],$positions[4])){$editorOpens+=Timed ('ScaleOpenEditor('+$ordinal+')');$editorCloses+=Timed 'ScaleCloseEditor()'}
  $history=$app.Eval('ScaleHistoryCount()');if($history -ne $startHistory){throw 'Live benchmark created group history'}
  if($app.Eval('UXGridPendingCount()') -ne 0){throw 'Pending cells left after benchmark'}
  $row=[ordered]@{Kind=$g.Kind;Size=$g.Size;GroupID=$g.GroupID;GroupOpen=Summary $opens;RowNavigation=Summary $moves;AutomaticCellSave=Summary $saves;EditorOpen=Summary $editorOpens;EditorClose=Summary $editorCloses;Raw=[ordered]@{GroupOpenMs=$opens;RowNavigationMs=$moves;CellSaveMs=$saves;EditorOpenMs=$editorOpens;EditorCloseMs=$editorCloses}}
  $results+=$row
  Write-Output ('Measured '+$g.Kind+' '+$g.Size+': open median '+$row.GroupOpen.MedianMs+' ms; row '+$row.RowNavigation.MedianMs+' ms; save '+$row.AutomaticCellSave.MedianMs+' ms')
  if(-not $app.Eval('UXCloseComparison()')){throw 'Final group close failed'}
 }
 $revisionDelta=$app.Eval('DCount("*","ReportRevision")')-$startRevisions;if($revisionDelta -ne 90){throw ('Expected 90 committed cell revisions, got '+$revisionDelta)}
 $evidence=[ordered]@{Date='2026-10-10';App='r102';Fixture=$fixture.Database;Scratch=$scratch;FixtureReports=525;Groups=9;Method='Single Access session on scratch copy; 3 opens, 15 row moves, 10 queued automatic saves and 3 editor opens/closes per group. Stopwatch around Access COM Eval; includes automation overhead. Native forms/events; no physical input or screen measurement.';HistoryRowsAdded=$app.Eval('ScaleHistoryCount()')-$startHistory;ReportRevisionsAdded=$revisionDelta;Results=$results;Limit='Sequential groups on this machine; synthetic observed drafts, all Promotable, no approvals; not production/printer/display or cold-start benchmark. Raw timings retained.'}
 $evidence|ConvertTo-Json -Depth 9|Set-Content (Join-Path $root 'planning/SCALE_R102_RESULTS.json') -Encoding UTF8
}finally{try{if($app.CurrentProject.AllForms.Item('frmUXComparison').IsLoaded){[void]$app.Eval('UXGridDiscardPending()')}}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
