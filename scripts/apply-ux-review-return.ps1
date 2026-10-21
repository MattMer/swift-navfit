# Focused saved-revision review and native Reports viewport return.
if(-not $script:swiftManagedBuild){foreach($module in @('modUXReviewLayout','modRouteTemplates','modUXReview','modUXReportReturn','modUX','modSwift','modMajorUI','modUXDashboard')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
NewScreen 'frmUXReview' 'Report review' 11200;$f=$script:swiftForm;$f.ScrollBars=2;$f.OnLoad='=UXLoadReview()'
foreach($name in @('txtReportID','txtGroupID','txtExpectedRevision','txtExpectedStamp')){[void](Control $f 109 $name 0 0 100 100);$script:swiftControl.Visible=$false}
[void](Control $f 100 'lblSaved' 400 1050 12600 650 '')
[void](Control $f 100 'lblTemplate' 400 1850 5900 300 'Review route template')
[void](Control $f 111 'cboTemplate' 400 2200 5900 450);$script:swiftControl.ColumnCount=2;$script:swiftControl.ColumnWidths='0;5900';$script:swiftControl.LimitToList=$true
[void](Control $f 104 'cmdStart' 6600 2200 6200 450 'Start review of current saved revision' '=UXReviewAction("Start")')
[void](Control $f 100 'lblRun' 400 2850 12600 300 'Current and previous review routes')
[void](Control $f 111 'cboRun' 400 3200 12600 450);$script:swiftControl.ColumnCount=2;$script:swiftControl.ColumnWidths='0;12600';$script:swiftControl.LimitToList=$true;$script:swiftControl.AfterUpdate='=UXReviewStatus()'
[void](Control $f 100 'lblStatus' 400 3850 12600 950 '')
[void](Control $f 110 'lstSteps' 400 4900 6000 1900);$script:swiftControl.ColumnCount=5;$script:swiftControl.ColumnHeads=$true;$script:swiftControl.ColumnWidths='650;3300;1800;0;0';$script:swiftControl.AfterUpdate='=UXReviewStep()'
[void](Control $f 109 'txtStep' 6600 4900 6200 1900);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2
[void](Control $f 100 'lblNote' 400 7000 12600 300 'Decision note / local approval note / reopening reason')
[void](Control $f 109 'txtNote' 400 7350 12600 800);$script:swiftControl.ScrollBars=2
$i=0
foreach($pair in @(@('Accept','Accept step'),@('Return','Return report'),@('Restart','Restart stale route'),@('Cancel','Cancel route'),@('Approve','Local approval'),@('Reopen','Reopen local approval'))){
[void](Control $f 104 ('cmd'+$pair[0]) (400+($i%3)*4300) (8350+[math]::Floor($i/3)*550) 4000 450 $pair[1] ('=UXReviewAction("'+$pair[0]+'")'));$i++
}
[void](Control $f 100 'lblUnavailable' 400 9500 12600 850 '')
[void](Control $f 104 'cmdBack' 400 10500 6000 450 'Back to workspace' '=UXCloseScreen("frmUXReview")')
[void](Control $f 104 'cmdRefresh' 6800 10500 6000 450 'Refresh saved review state' '=UXLoadReview()')
$i=0;foreach($name in @('cboTemplate','cmdStart','cboRun','lstSteps','txtStep','txtNote','cmdAccept','cmdReturn','cmdRestart','cmdCancel','cmdApprove','cmdReopen','cmdBack','cmdRefresh')){$f.Controls($name).TabIndex=$i;$i++}
foreach($c in $f.Controls){if($c.Visible){$c.Tag='RV:'+$c.Left+':'+$c.Width+':'+$c.Top+':'+$c.Height}}
$f.OnResize='=UXReviewLayout()'
SaveScreen $f 'frmUXReview'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor');$f.OnClose='=UXQueueReportReturn()';SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
[void](Control $f 104 'uxFocusedReview' 2850 6300 10450 500 'Review selected member - saved revision' '=UXGroupReview()');$script:swiftControl.Tag='WBX:Review'
SaveScreen $f 'frmMajorWorkbench'



$app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary');$f.OnTimer='=UXFinishReportReturn()';$f.TimerInterval=0;SaveScreen $f 'frmLibrary'
