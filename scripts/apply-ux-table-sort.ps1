foreach($m in @('modTableSort','modUXTraitGrid','modUXSeniorTable','modUXViewport')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$coverage=@()
foreach($screen in @($app.CurrentProject.AllForms|ForEach-Object {$_.Name})){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 $lists=@($f.Controls|Where-Object {$_.ControlType -eq 110 -and $_.ColumnCount -gt 1})
 $custom=$false
 foreach($c in $lists){
  $widths=@($c.ColumnWidths -split ';');$visible=0
  for($i=0;$i -lt $c.ColumnCount;$i++){if($i -ge $widths.Count -or $widths[$i] -eq '' -or [double]$widths[$i] -gt 0){$visible++}}
  if($visible -lt 2){continue}
  if($c.ColumnHeads){
   $f.HasModule=$true
   $module=$f.Module
   if($c.OnMouseDown -eq '[Event Procedure]'){$line=$module.ProcStartLine($c.Name+'_MouseDown',0)}else{$line=$module.CreateEventProc('MouseDown',$c.Name)}
   $code='    If UXSortNativeHeader(Me.Name, "'+$c.Name+'", X, Y) Then Exit Sub'
   if($module.Lines(1,$module.CountOfLines) -notlike ('*'+$code+'*')){$module.InsertLines($line+1,$code)}
   $c.OnMouseDown='[Event Procedure]'
   $coverage+=@{Form=$screen;Control=$c.Name;Method='Native header click'}
  }else{
   $custom=$true
   #Reserve one header strip without changing the bottom edge or list contents.
   $already=$false
   try{$already=($f.Controls('sort_'+$c.Name+'_0').Tag -like 'SORT:*')}catch{}
   $top=$c.Top
   if($already){$top-=350}else{$c.Top=[int]($c.Top+350);$c.Height=[math]::Max(600,$c.Height-350)}
   for($i=0;$i -lt $c.ColumnCount;$i++){
    $name='sort_'+$c.Name+'_'+$i
    try{$app.DeleteControl($f.Name,$name)}catch{}
    $h=$app.CreateControl($f.Name,100,0,'','',$c.Left,$top,800,300);$h.Name=$name;$h.Caption='Sort';$h.FontName='Segoe UI';$h.FontSize=9;$h.FontUnderline=$true
    $h.Tag='SORT:'+$c.Name+':'+$i;$h.OnClick='=UXSortList("'+$screen+'","'+$c.Name+'",'+$i+')';$h.ForeColor=6177536
   }
   $coverage+=@{Form=$screen;Control=$c.Name;Method='Custom clickable headings'}
  }
 }
 if($custom){
  $original=$f.OnActivate
  if($original -like '=UXSortActivate*'){
   #Already wrapped; retain its original activation expression.
  }elseif($original -eq '[Event Procedure]'){
   $module=$f.Module;$line=$module.ProcStartLine('Form_Activate',0);$module.InsertLines($line+1,'    UXSortHeaders Me')
  }else{
   try{$app.DeleteControl($f.Name,'uxSortActivateOriginal')}catch{}
   $h=$app.CreateControl($f.Name,109,0,'','',0,0,100,100);$h.Name='uxSortActivateOriginal';$h.Visible=$false;$h.DefaultValue='"'+$original.Replace('"','""')+'"'
   $f.OnActivate='=UXSortActivate("'+$screen+'")'
  }
 }
 #Saving first then closing with acSaveNo can discard new class procedures.
 $app.DoCmd.Close(2,$f.Name,1)
}
$app.DoCmd.OpenForm('frmUXComparison',1);$grid=[string]$app.Forms.Item('frmUXComparison').Controls('uxTraitGrid').SourceObject;$app.DoCmd.Close(2,'frmUXComparison',2)
$grid=$grid -replace '^Form\.','';$app.DoCmd.OpenForm($grid,1);$f=$app.Forms.Item($grid)
foreach($h in @($f.Controls|Where-Object {$_.ControlType -eq 100 -and $_.Section -eq 1})){
 $best=$null;$distance=999999
 foreach($c in $f.Controls){if($c.ControlType -in @(109,111) -and $c.Section -eq 0 -and $c.ControlSource -and $c.Visible){$d=[math]::Abs($c.Left-$h.Left);if($d -lt $distance){$distance=$d;$best=$c}}}
 if($null -ne $best){$h.OnClick='=UXSortGrid("'+$best.ControlSource+'")';$h.ControlTipText='Click to sort; click again to reverse';$coverage+=@{Form=$grid;Control=$h.Name;Method='Editable grid header';Field=$best.ControlSource}}
}
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmUXRSCATable',1);$f=$app.Forms.Item('frmUXRSCATable')
foreach($key in @('grade','avg','count','date')){$f.Controls('head'+$key).OnClick='=RSTableSort("'+$key+'")';$f.Controls('head'+$key).ControlTipText='Click to sort; click again to reverse';$coverage+=@{Form=$f.Name;Control=('head'+$key);Method='RSCA table header'}}
$app.DoCmd.Close(2,$f.Name,1)
$coverage|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root 'build/r115-sort-coverage.json')
