# Final presentation pass. All controls retain their service-facing identities.
# Coordinates are twips; the 920-point workspace fits a laptop at 100% scaling.
$uxDestinations=@(@('Overview','frmHome'),@('Reports','frmLibrary'),@('Summary groups','frmUXGroups'),@('My Work','frmMyWork'),@('Resources','frmUXResources'),@('Settings & Help','frmUXSettings'))
$uxPages=@('frmHome','frmLibrary','frmEditor','frmUXGroups','frmMyWork','frmUXResources','frmUXSettings','frmMajorWorkbench','frmUXBulk','frmUXComparison','frmUXSeniors','frmUXSeniorEditor','frmUXGroupRSCA')
foreach($screenName in $uxPages){
    $app.DoCmd.OpenForm($screenName,1);$f=$app.Forms.Item($screenName)
    # Reserve the same navigation rail on every destination. Editor section
    # navigation stays inside the content area, distinct from app navigation.
    foreach($c in $f.Controls){
        if($c.Name -match '^uxNav\d+$' -or $c.Name -eq 'uxContext'){continue}
        if($c.Left -gt 100 -or $c.Width -gt 100){
            $c.Left=[int](2550+$c.Left*0.79);$c.Width=[int]([math]::Max(200,$c.Width*0.79))
        }
        try{$c.FontName='Segoe UI'}catch{}
        if($c.ControlType -eq 104){$c.UseTheme=$false;$c.BackColor=16777215;$c.ForeColor=3946554;$c.BorderColor=14012862}
        if($c.ControlType -in @(109,110,111)){$c.BackColor=16777215;$c.ForeColor=2368548;$c.BorderColor=14012862}
    }
    $f.Width=[int](13800);$f.Section(0).BackColor=16381429
    [void](Control $f 101 'uxRail' 0 0 2450 $f.Section(0).Height);$script:swiftControl.BackStyle=0;$script:swiftControl.BackColor=16777215;$script:swiftControl.BorderColor=15130328
    [void](Control $f 100 'uxBrand' 220 250 2100 500 'Swift NAVFIT');$script:swiftControl.FontSize=17;$script:swiftControl.ForeColor=6897174
    [void](Control $f 100 'uxNavCaption' 220 1000 2100 300 'WORKSPACE');$script:swiftControl.FontSize=9;$script:swiftControl.ForeColor=7955800
    for($i=0;$i -lt $uxDestinations.Count;$i++){
        $name='uxNav'+$i
        if(@($f.Controls|Where-Object Name -eq $name).Count -eq 0){[void](Control $f 104 $name 180 (1450+$i*650) 2100 550 $uxDestinations[$i][0] ('=UXNavigate("'+$uxDestinations[$i][1]+'")'))}
        $c=$f.Controls($name);$c.Left=[int](180);$c.Top=[int](1450+$i*650);$c.Width=[int](2100);$c.Height=[int](550);$c.Visible=$true;$c.UseTheme=$false;$c.BackColor=16777215;$c.ForeColor=5916729;$c.BorderStyle=0
        if($screenName -eq $uxDestinations[$i][1] -or ($i -eq 1 -and $screenName -eq 'frmEditor') -or ($i -eq 2 -and $screenName -in @('frmMajorWorkbench','frmUXBulk','frmUXComparison','frmUXGroupRSCA')) -or ($i -eq 4 -and $screenName -in @('frmUXSeniors','frmUXSeniorEditor'))){$c.BackColor=16445671;$c.ForeColor=9785094;$c.FontBold=$true}
    }
    if(@($f.Controls|Where-Object Name -eq 'uxContext').Count -eq 0){[void](Control $f 100 'uxContext' 2850 300 10300 400 'Swift NAVFIT | Local workspace')}
    $f.Controls('uxContext').Left=[int](2850);$f.Controls('uxContext').Top=[int](300);$f.Controls('uxContext').Width=[int](10300);$f.Controls('uxContext').Height=[int](400)
    # Preserve destination OnActivate handlers; context refresh also runs when
    # navigation opens a destination through its existing load routines.
    if(-not $f.OnActivate){$f.OnActivate='=UXLoadContext()'}
    if($screenName -eq 'frmHome'){
        foreach($name in @('cmdBackup','cmdOperations','cmdSupport','cmdSnippets')){$f.Controls($name).Visible=$false}
        $f.Controls('lblHeading').Caption='Overview'
        $cards=@(@('cmdLibrary','Reports','Create, find and continue your reports.'),@('cmdMajor','Summary groups','Manage members and compare trait marks.'),@('cmdMyWork','My Work','Continue assigned reviews.'),@('cmdRecovery','Recovery','Review available recovery checkpoints.'))
        for($i=0;$i -lt $cards.Count;$i++){
            $x=2850+($i%2)*5300;$y=2200+[math]::Floor($i/2)*2100
            [void](Control $f 101 ('uxCard'+$i) $x $y 4950 1800);$script:swiftControl.BackStyle=0;$script:swiftControl.BackColor=16777215;$script:swiftControl.BorderColor=15130328
            $c=$f.Controls($cards[$i][0]);$c.Left=[int]($x+250);$c.Top=[int]($y+250);$c.Width=[int](4450);$c.Height=[int](600);$c.Caption=$cards[$i][1];$c.FontSize=16
            [void](Control $f 100 ('uxCardHint'+$i) ($x+250) ($y+1000) 4450 550 $cards[$i][2]);$script:swiftControl.ForeColor=7955800
        }
        $f.Controls('cmdImportReview').Left=[int](2850);$f.Controls('cmdImportReview').Top=[int](6800);$f.Controls('cmdImportReview').Width=[int](4950);$f.Controls('cmdImportReview').Caption='Import NAVFIT database'
        $f.Controls('cmdPDF').Left=[int](8150);$f.Controls('cmdPDF').Top=[int](6800);$f.Controls('cmdPDF').Width=[int](4950);$f.Controls('cmdPDF').Caption='Import structured PDF'
    }
    if($screenName -eq 'frmEditor'){
        $f.Controls('uxReport').Width=[int](4500)
        $c=$f.Controls('cboKind');$c.Visible=$true;$c.Left=[int](11000);$c.Top=[int](1000);$c.Width=[int](2150);$c.Height=[int](450)
        [void](Control $f 100 'uxTypeCaption' 9300 1050 1550 350 'Report type')
    }
    if($screenName -eq 'frmLibrary'){
        $f.Controls('lblHeading').Caption='Reports'
        foreach($name in @('lstFolders','lblFolders','cmdHome','cmdFolderHelp','lblGroupScope')){$f.Controls($name).Visible=$false}
        # A folder is a filter, rather than a second full-height navigation tree.
        [void](Control $f 111 'uxFolderFilter' 2850 1650 2900 450);$c=$script:swiftControl;$c.RowSource="SELECT 0 AS ID, 'All reports' AS Label FROM AppMetadata UNION ALL SELECT -1, 'Unassigned' FROM AppMetadata UNION ALL SELECT FolderID, FolderPath([FolderID]) FROM Folder ORDER BY Label";$c.ColumnCount=2;$c.ColumnWidths='0;2800';$c.AfterUpdate='=UXFolderFilterChanged()'
        [void](Control $f 100 'uxFolderCaption' 2850 1250 2900 300 'Folder')
        $f.Controls('txtSearch').Left=[int](6100);$f.Controls('txtSearch').Top=[int](1650);$f.Controls('txtSearch').Width=[int](3400)
        $f.Controls('lblSearch').Left=[int](6100);$f.Controls('lblSearch').Top=[int](1250);$f.Controls('lblSearch').Width=[int](3400);$f.Controls('lblSearch').Caption='Search reports'
        $f.Controls('cboArchive').Left=[int](9850);$f.Controls('cboArchive').Top=[int](1650);$f.Controls('cboArchive').Width=[int](1450)
        $f.Controls('cmdFind').Left=[int](11500);$f.Controls('cmdFind').Top=[int](1650);$f.Controls('cmdFind').Width=[int](1800)
        $f.Controls('lstReports').Left=[int](2850);$f.Controls('lstReports').Top=[int](2450);$f.Controls('lstReports').Width=[int](10450);$f.Controls('lstReports').Height=[int](3000)
        $f.Controls('lstReports').ColumnWidths='0;1200;3000;2700;1300;900'
        $f.Controls('cmdNew').Caption='New report';$f.Controls('cmdOpen').Caption='Open report'
    }
    $tabIndex=0
    foreach($c in @($f.Controls | Where-Object {$_.ControlType -in @(104,106,109,110,111)} | Sort-Object @{Expression={if($_.Name -match '^uxNav\d+$'){0}else{1}}},Top,Left)){
        $c.TabIndex=[int]$tabIndex;$tabIndex++
    }
    SaveScreen $f $screenName
}
