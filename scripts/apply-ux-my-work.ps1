# Unified saved-review inbox and native review findings. No schema changes.
if(-not $script:swiftManagedBuild){foreach($module in @('modUXReadRows','modUXHelp','modUXMyWork','modUXReportReturn','modRouting','modReviewComments','modUXReview','modUXDashboard','modUXSurface','modLibraryViews','modUX')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
$app.DoCmd.OpenForm('frmMyWork',1);$f=$app.Forms.Item('frmMyWork')
[void](Control $f 110 'lstWork' 2850 3000 10450 1900);$c=$script:swiftControl;$c.ColumnCount=11;$c.ColumnHeads=$true;$c.ColumnWidths='0;0;0;0;0;3000;1600;1700;1100;1500;2200';$c.AfterUpdate='=UXWorkSelection()';$c.OnDblClick='=UXWorkOpen()';$c.Tag='MWX:body'
[void](Control $f 100 'uxWorkStatus' 2850 2450 8000 450 '');$script:swiftControl.Tag='MWX:body'
[void](Control $f 104 'cmdOpenWork' 2850 8100 10450 450 'Open selected work' '=UXWorkOpen()');$script:swiftControl.Tag='MWX:body'
$f.Controls('lstTasks').Visible=$false
$f.Controls('txtHistory').Top=5850;$f.Controls('txtHistory').Height=1300;$f.Controls('txtHistory').ScrollBars=2
$f.Controls('lblNotes').Top=7270;$f.Controls('txtNotes').Top=7650;$f.Controls('txtNotes').Height=350
$f.Controls('chkAll').Top=2500;$f.Controls('lblAll').Top=2500
$f.Controls('uxWorkStatus').Top=2900;$f.Controls('uxWorkStatus').Height=400
$f.Controls('lstWork').Top=3400;$f.Controls('lstWork').Height=2300
$f.Controls('lblAll').Caption='Show all contributors and completed history'
$f.Controls('cmdDraft').OnClick='=UXWorkOpen(True)';$f.Controls('cmdDraft').Top=8700
foreach($name in @('cmdAccept','cmdReturn','cmdCancel')){$f.Controls($name).Top=8700}
$f.Controls('uxWorkTab0').Caption='My review work'
$f.OnActivate='=UXWorkChanged()'
$f.OnResize='=UXWorkLayout()'
$i=0;foreach($name in @('uxWorkTab0','uxWorkTab1','uxWorkTab2','chkAll','cmdRefresh','lstWork','txtHistory','txtNotes','cmdOpenWork','cmdDraft','cmdAccept','cmdReturn','cmdCancel')){$f.Controls($name).TabIndex=$i;$i++}
SaveScreen $f 'frmMyWork'

$app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary')
$f.Controls('uxFolderFilter').LimitToList=$true
# Keep headings above the filter row, rather than beneath its labels.
$f.Controls('lblHeading').Top=800;$f.Controls('lblHeading').Height=450
foreach($name in @('uxFolderCaption','lblSearch')){$f.Controls($name).Top=1400}
foreach($name in @('uxFolderFilter','txtSearch','cboArchive','cmdFind')){$f.Controls($name).Top=1750}
$f.Controls('lstReports').Height=2300
SaveScreen $f 'frmLibrary'
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups')
$f.Controls('uxFilterLabel').Top=2350;$f.Controls('uxArchiveFilter').Top=2300
$f.Controls('lstGroups').Top=2950;$f.Controls('lstGroups').Height=2900
SaveScreen $f 'frmUXGroups'
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison')
$f.Controls('lblHeading').Top=800;$f.Controls('lblHeading').Height=450
SaveScreen $f 'frmUXComparison'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
$f.Controls('cmdSave').Caption='Save draft';$f.Controls('cmdSave').ForeColor=5916729;$f.Controls('cmdSave').UseTheme=$false;$f.Controls('cmdSave').BackColor=16445671
$f.Controls('lstRevisions').ColumnHeads=$true
SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench');$f.Controls('lstReports').ColumnHeads=$true;SaveScreen $f 'frmMajorWorkbench'
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome');$f.OnActivate='=UXDashboardRefresh()';SaveScreen $f 'frmHome'
$app.DoCmd.OpenForm('frmUXSettings',1);$f=$app.Forms.Item('frmUXSettings');$f.Controls('lblHeading').Caption='Settings && Help';SaveScreen $f 'frmUXSettings'
$app.DoCmd.OpenForm('frmUXReview',1);$f=$app.Forms.Item('frmUXReview')
[void](Control $f 104 'cmdMyWork' 400 11100 12600 450 'Back to My Work' '=UXWorkBack()');$script:swiftControl.Tag='RV:400:12600:11100:450';$f.Section(0).Height=12000;$f.OnClose='=UXWorkChanged()'
SaveScreen $f 'frmUXReview'

# Bring high-frequency supporting screens into the same workspace shell.
$destinations=@(@('Overview','frmHome'),@('Reports','frmLibrary'),@('Summary groups','frmUXGroups'),@('My Work','frmMyWork'),@('Resources','frmUXResources'),@('Settings & Help','frmUXSettings'))
foreach($screenName in @('frmReviewComments','frmSnippets','frmRecovery','frmUXImportChoose')){
    $app.DoCmd.OpenForm($screenName,1);$f=$app.Forms.Item($screenName)
    foreach($c in $f.Controls){
        if($c.Visible){$c.Left=[int](2850+($c.Left-400)*0.84);$c.Width=[int]([Math]::Max(200,$c.Width*0.84));$c.Top=[int]($c.Top+600)}
        try{$c.FontName='Segoe UI'}catch{}
        if($c.ControlType -eq 104){$c.UseTheme=$false;$c.BackColor=16777215;$c.ForeColor=5916729}
    }
    [void](Control $f 101 'uxRail' 0 0 2450 10500);$script:swiftControl.BackStyle=0
    [void](Control $f 100 'uxBrand' 220 250 2100 500 'Swift NAVFIT');$script:swiftControl.FontSize=17
    [void](Control $f 100 'uxNavCaption' 220 1000 2100 300 'WORKSPACE')
    for($i=0;$i -lt 6;$i++){[void](Control $f 104 ('uxNav'+$i) 180 (1450+$i*650) 2100 550 $destinations[$i][0].Replace('&','&&') ('=UXNavigate("'+$destinations[$i][1]+'")'));$script:swiftControl.UseTheme=$false;$script:swiftControl.ForeColor=5916729}
    [void](Control $f 100 'uxContext' 2850 200 10300 400 'Swift NAVFIT | Local workspace')
    [void](Control $f 111 'uxCompactNav' 250 150 2500 400);$c=$script:swiftControl;$c.Visible=$false;$c.RowSourceType='Value List';$c.RowSource='frmHome;Overview;frmLibrary;Reports;frmUXGroups;Summary groups;frmMyWork;My Work;frmUXResources;Resources;frmUXSettings;Settings & Help';$c.ColumnCount=2;$c.ColumnWidths='0;2500';$c.LimitToList=$true;$c.AfterUpdate='=UXCompactNavigate("'+$screenName+'")'
    $f.Width=13800;$f.Section(0).BackColor=16381429;$f.OnResize='=UXResizeWorkspace("'+$screenName+'")';$f.OnActivate='=UXLoadContext()'
    if($screenName -eq 'frmReviewComments'){
        $f.Controls('lstThreads').ColumnHeads=$true
        [void](Control $f 104 'cmdMyWork' 2850 14500 10450 500 'Back to My Work' '=UXWorkBack()');$f.OnClose='=UXWorkChanged()'
    }
    if($screenName -eq 'frmSnippets'){$f.Controls('cmdClose').OnClick='=UXCloseResource("frmSnippets")'}
    SaveScreen $f $screenName
}
