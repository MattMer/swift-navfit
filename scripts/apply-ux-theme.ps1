# Final presentation pass: Access OLE colors use R + G*256 + B*65536.
$theme = @{ Canvas=15397112; Paper=16777215; Ink=2501670; Border=9682379; Peach=7254527; Ochre=5948108; Yellow=12121599; Khaki=9682379 }
$names=@($app.CurrentProject.AllForms | ForEach-Object {$_.Name})
$controls=0
foreach($name in $names){
 $app.DoCmd.OpenForm($name,1)
 $f=$app.Forms.Item($name)
 $f.Section(0).BackColor=$theme.Canvas
 foreach($c in $f.Controls){
  switch([int]$c.ControlType){
   100 { $c.ForeColor=$theme.Ink; if($c.Name -like 'err*' -or $c.Name -eq 'uxError'){$c.ForeColor=2893478} }
   104 { $c.UseTheme=$false; $c.BackColor=$theme.Peach; $c.ForeColor=$theme.Ink; $c.BorderColor=$theme.Ochre
         if($c.Name -in @('cmdSave','cmdApply','cmdImport','cmdExport','cmdCreate')){$c.BackColor=$theme.Ochre}
         if($c.Name -like 'uxNav*' -or $c.Name -like 'uxSection*'){$c.BackColor=if($c.FontBold){$theme.Ochre}else{$theme.Yellow}} }
   109 { $c.BackColor=$theme.Paper; $c.ForeColor=$theme.Ink; $c.BorderColor=$theme.Border }
   110 { $c.BackColor=$theme.Paper; $c.ForeColor=$theme.Ink; $c.BorderColor=$theme.Border }
   111 { $c.BackColor=$theme.Paper; $c.ForeColor=$theme.Ink; $c.BorderColor=$theme.Border }
   106 { } # Checkbox captions are separate label controls.
   101 { $c.BorderColor=$theme.Border }
  }
  $controls++
 }
 $app.DoCmd.Close(2,$name,1)
}
Write-Output ('Siltstone theme applied: '+$names.Count+' forms, '+$controls+' controls.')
