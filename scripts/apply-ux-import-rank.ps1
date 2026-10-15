foreach($m in @('modNAVFITAdapter','modPDFImport','modUXNativeGroups','modUXNAVFIT','modUXImportFlow','modUXDirectory','modUXGroupPages','modUXHelp')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
foreach($entry in @(@('frmPDFImport','cboCentury','lblCentury'),@('frmUXNativeGroups','cboCentury','lblCentury'),@('frmImportReview','cboNAVFITCentury','lblNAVFITCentury'),@('frmNAVFITExchange','cboReturnCentury','lblReturnCentury'))){
 $app.DoCmd.OpenForm($entry[0],1);$f=$app.Forms.Item($entry[0])
 foreach($name in $entry[1..2]){try{$app.DeleteControl($f.Name,$name)}catch{}}
 $app.DoCmd.Close(2,$f.Name,1)
}
foreach($screen in @('frmUXGroupReport','frmUXDirectoryBatch')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 foreach($name in @('lblRankFilter','cboRankFilter')){try{$app.DeleteControl($f.Name,$name)}catch{}}
 if($screen -eq 'frmUXGroupReport'){
  foreach($c in $f.Controls){if($c.Top -ge 1650){$c.Top=[int]($c.Top+850)}}
  $labelTop=1650;$controlTop=2050;$left=400;$width=11000
  $handler='=UXGroupMemberFilter()'
 }else{
  foreach($c in $f.Controls){if($c.Top -ge 2250){$c.Top=[int]($c.Top+450)}}
  $labelTop=2130;$controlTop=2050;$left=400;$width=12600
  $handler='=UXDirectoryRankFilter()'
 }
 $c=$app.CreateControl($f.Name,100,0,'','',$left,$labelTop,2000,350);$c.Name='lblRankFilter';$c.Caption='Filter by rank/rate';$c.FontName='Segoe UI';$c.FontSize=10
 if($screen -eq 'frmUXDirectoryBatch'){$left=2500;$width=10500}
 $c=$app.CreateControl($f.Name,111,0,'','',$left,$controlTop,$width,400);$c.Name='cboRankFilter';$c.RowSourceType='Value List';$c.LimitToList=$true;$c.AfterUpdate=$handler;$c.FontName='Segoe UI';$c.FontSize=10
 $c.ControlTipText='Choose a paygrade (such as E6) or a specific rank/rate. All ranks/rates shows every active member.'
 $order=if($screen -eq 'frmUXGroupReport'){@('cboRankFilter','cboMember','txtName','txtOfficial','txtRank','txtSSN','txtFrom','txtTo','cmdCreate','cmdCancel')}else{@('cboGroup','txtFrom','txtTo','cboRankFilter','lstMembers','cmdPreview','txtPreview','cmdApply','cmdClose')}
 for($i=0;$i -lt $order.Count;$i++){$f.Controls($order[$i]).TabIndex=$i}
 $f.Section(0).Height=[math]::Max($f.Section(0).Height,($f.Controls($(if($screen -eq 'frmUXGroupReport'){'cmdCreate'}else{'cmdApply'})).Top+1000))
 $app.DoCmd.Close(2,$f.Name,1)
}
