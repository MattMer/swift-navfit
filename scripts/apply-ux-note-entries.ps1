foreach($m in @('modUXNotes','modUXSurface','modRouting')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmUXEditorFields',1);$f=$app.Forms.Item('frmUXEditorFields')
[void](Control $f 110 'uxNoteList' 200 200 8000 1750);$c=$script:swiftControl;$c.ColumnCount=3;$c.ColumnHeads=$true;$c.OnClick='=UXSelectReportNote()'
[void](Control $f 109 'uxNoteDetails' 200 2100 8000 1650);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2
[void](Control $f 104 'uxAddNote' 200 3900 8000 450 'Add contributor note' '=UXOpenNote(False)')
[void](Control $f 110 'uxNoteDrafts' 200 4500 8000 1000);$c=$script:swiftControl;$c.ColumnCount=2
[void](Control $f 104 'uxRecoverNote' 200 5650 8000 450 'Recover selected unposted note' '=UXOpenNote(True)')
SaveScreen $f 'frmUXEditorFields'
NewScreen 'frmUXNoteComposer' 'Contributor note' 6500;$f=$script:swiftForm;$f.PopUp=$true;$f.ScrollBars=2
[void](Control $f 100 'lblContributor' 400 1000 9600 400 '')
[void](Control $f 109 'txtNote' 400 1650 9600 2900);$c=$script:swiftControl;$c.ScrollBars=2;$c.AfterUpdate='=UXNoteCheckpoint()'
[void](Control $f 100 'lblStatus' 400 4800 9600 750 '')
[void](Control $f 104 'cmdPost' 400 5750 3000 500 'Post note' '=UXPostNote()')
[void](Control $f 104 'cmdDiscard' 7000 5750 3000 500 'Discard this note' '=UXDiscardNote()')
[void](Control $f 104 'cmdKeep' 3700 5750 3000 500 'Close and keep checkpoint' '=UXKeepNoteCheckpoint()')
$f.OnLoad='=UXLoadNoteComposer()';$f.OnTimer='=UXNoteCheckpoint()';$f.TimerInterval=30000;$f.OnUnload='=UXNoteUnload()';$f.OnClose='=UXNoteClosed()'
SaveScreen $f 'frmUXNoteComposer'


$app.LoadFromText(5,'modUXDashboard',(Join-Path $root 'src/modUXDashboard.bas'))
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome')
[void](Control $f 104 'uxNoteRecovery' 7800 2090 5500 420 'Unposted notes' '=UXOpenNoteRecoveries()')
SaveScreen $f 'frmHome'
NewScreen 'frmUXNoteRecovery' 'My unposted notes' 6700;$f=$script:swiftForm
[void](Control $f 110 'lstDrafts' 400 1250 12600 2900);$c=$script:swiftControl;$c.ColumnCount=3;$c.ColumnWidths='0;7800;4300';$c.OnDblClick='=UXResumeSelectedNote()'
[void](Control $f 100 'lblStatus' 400 4400 12600 850 'Select a note checkpoint to resume it alongside its report. Archived reports must be recovered before posting.')
[void](Control $f 104 'cmdResume' 400 5550 6100 500 'Resume selected note' '=UXResumeSelectedNote()')
[void](Control $f 104 'cmdBack' 6800 5550 6200 500 'Back to Overview' '=UXCloseResource("frmUXNoteRecovery")')
SaveScreen $f 'frmUXNoteRecovery'
