if(-not $script:swiftManagedBuild){foreach($m in @('modUXNotes','modUXSurface','modRouting')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXDashboard',(Join-Path $root 'src/modUXDashboard.bas'))}
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome')
[void](Control $f 104 'uxNoteRecovery' 7800 2090 5500 420 'Unposted notes' '=UXOpenNoteRecoveries()')
SaveScreen $f 'frmHome'
NewScreen 'frmUXNoteRecovery' 'My unposted notes' 6700;$f=$script:swiftForm
[void](Control $f 110 'lstDrafts' 400 1250 12600 2900);$c=$script:swiftControl;$c.ColumnCount=3;$c.ColumnWidths='0;7800;4300';$c.OnDblClick='=UXResumeSelectedNote()'
[void](Control $f 100 'lblStatus' 400 4400 12600 850 'Select a note checkpoint to resume it alongside its report. Archived reports must be recovered before posting.')
[void](Control $f 104 'cmdResume' 400 5550 6100 500 'Resume selected note' '=UXResumeSelectedNote()')
[void](Control $f 104 'cmdBack' 6800 5550 6200 500 'Back to Overview' '=UXCloseResource("frmUXNoteRecovery")')
SaveScreen $f 'frmUXNoteRecovery'
