param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/note-shutdown-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.LoadFromText(5,'UXNoteEntriesHarness',(Join-Path $root 'tests/UXNoteEntriesHarness.bas'));$app.DoCmd.RunCommand(126);[void]$app.Eval('StageNoteShutdown()');$app.Quit(2)}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($scratch,$false,$true)
try{$r=$db.OpenRecordset('SELECT NoteText FROM ReportNoteDraft',2);if($r.EOF -or $r.Fields(0).Value -cne 'RAW UNSUBMITTED NOTE AT ACCESS SHUTDOWN'){throw 'Raw note did not survive shutdown'};$r.Close();$r=$db.OpenRecordset('SELECT Count(*) FROM ReportNote',2);if($r.Fields(0).Value -ne 0){throw 'Shutdown posted a note without the contributor action'};$r.Close()}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
Write-Output 'PASS: Access shutdown retains raw unposted note text without posting it'
