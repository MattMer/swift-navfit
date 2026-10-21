. (Join-Path $PSScriptRoot 'access-properties.ps1')
if(-not $script:swiftManagedBuild){foreach($m in @('modMigration','modWorkspaceImport','modReportPageMap','modReportPages','modUX','modUXSurface','modUXValidation','modUXGroupLifecycle','modUXNotes','modUXHelp')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
function PageControl($f,$type,$name,$x,$y,$w,$h,$caption='',$event=''){
 $c=$app.CreateControl($f.Name,$type,0,'','',$x,$y,$w,$h);$c.Name=$name
 if($type -in @(100,104)){$c.Caption=$caption}
 $c.FontName='Segoe UI';$c.FontSize=10
 if($event){$c.OnClick=$event}
 $script:pageControl=$c
}
$app.LoadFromText(2,'frmUXWorkspaceImport',(Join-Path $root 'src/BlankForm.txt'))
$app.DoCmd.OpenForm('frmUXWorkspaceImport',1);$f=$app.Forms.Item('frmUXWorkspaceImport');$f.Caption='Import previous Swift NAVFIT workspace';$f.Width=13000;$f.Section(0).Height=9800;$f.ScrollBars=2
PageControl $f 100 'lblHeading' 350 250 12200 650 'Import previous Swift NAVFIT workspace'
PageControl $f 109 'txtPreview' 350 1200 12200 5900;$script:pageControl.Locked=$true;$script:pageControl.ScrollBars=2
PageControl $f 104 'cmdApply' 350 7400 6000 500 'Import all workspace data' '=UXApplySwiftWorkspace()'
PageControl $f 104 'cmdClose' 6600 7400 5950 500 'Close' '=UXCloseScreen("frmUXWorkspaceImport")'
PageControl $f 100 'lblStatus' 350 8150 12200 1100 ''
$app.DoCmd.Save(2,$f.Name);$app.DoCmd.Close(2,$f.Name,2)
foreach($screen in @('frmUXSettings','frmSetup')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen);$bottom=0
 foreach($c in $f.Controls){if($c.Visible -and ($screen -eq 'frmSetup' -or $c.Left -ge 2450)){$bottom=[math]::Max($bottom,$c.Top+$c.Height)}}
 try{$app.DeleteControl($f.Name,'cmdImportSwift')}catch{}
 $left=if($screen -eq 'frmSetup'){400}else{2850}
 PageControl $f 104 'cmdImportSwift' $left ($bottom+250) 9950 550 'Import data from a previous Swift NAVFIT version' '=UXChooseSwiftWorkspace()'
 SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,[int]($bottom+1100)));$f.ScrollBars=2
 $app.DoCmd.Save(2,$f.Name);$app.DoCmd.Close(2,$f.Name,2)
}
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
#Hide old section/navigation chrome; canonical buffers and backend identities stay.
foreach($c in $f.Controls){if($c.Name -like 'uxSection?' -or $c.Name -in @('uxCompactSection','uxPage','uxPrevious','uxNext','lblHeading','lblImportOrigin')){$c.Visible=$false}}
foreach($item in @(@('uxFrontPage','Front Page','=UXSelectSection(0)'),@('uxBackPage','Back Page','=UXSelectSection(5)'),@('uxPageNotes','Draft notes','=UXSelectSection(7)'),@('uxPageReview','Review','=UXSelectSection(8)'),@('uxPageHistory','History','=UXSelectSection(9)'),@('uxPageValidation','Validation issues','=UXOpenValidation(False)'))){
 try{$app.DeleteControl($f.Name,$item[0])}catch{}
 PageControl $f 104 $item[0] 250 2850 2200 400 $item[1] $item[2]
}
$f.Width=15000;$f.ScrollBars=3
$app.DoCmd.Save(2,$f.Name);$app.DoCmd.Close(2,$f.Name,2)
