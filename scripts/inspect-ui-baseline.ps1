param([Parameter(Mandatory=$true)][string]$DatabasePath,[Parameter(Mandatory=$true)][string]$EvidencePath)
$ErrorActionPreference='Stop'
$DatabasePath=[IO.Path]::GetFullPath($DatabasePath);$EvidencePath=[IO.Path]::GetFullPath($EvidencePath)
[void](New-Item -ItemType Directory -Force -Path $EvidencePath)
$before=(Get-FileHash -LiteralPath $DatabasePath -Algorithm SHA256).Hash
$scratch=Join-Path $EvidencePath 'inspection.accdb';Copy-Item -LiteralPath $DatabasePath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=3;$app.OpenCurrentDatabase($scratch,$true)
 foreach($n in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$n,2)}
 function Properties($object){
  $result=[ordered]@{}
  foreach($p in $object.Properties){try{$v=$p.Value;if($null -eq $v -or $v -is [string] -or $v.GetType().IsPrimitive -or $v -is [decimal] -or $v -is [datetime]){$result[$p.Name]=$v}}catch{}}
  return $result
 }
 $forms=@()
 foreach($n in @($app.CurrentProject.AllForms|ForEach-Object {$_.Name})){
  $app.SaveAsText(2,$n,(Join-Path $EvidencePath ($n+'.form.txt')))
  $app.DoCmd.OpenForm($n,1);$f=$app.Forms.Item($n);$controls=@()
  foreach($c in $f.Controls){$controls+=@{name=$c.Name;type=[int]$c.ControlType;properties=(Properties $c)}}
  $forms+=@{name=$n;properties=(Properties $f);controls=$controls}
  $app.DoCmd.Close(2,$n,2)
 }
 $modules=@();foreach($n in @($app.CurrentProject.AllModules|ForEach-Object {$_.Name})){$app.SaveAsText(5,$n,(Join-Path $EvidencePath ($n+'.bas.txt')));$modules+=$n}
 $macros=@();foreach($n in @($app.CurrentProject.AllMacros|ForEach-Object {$_.Name})){$app.SaveAsText(4,$n,(Join-Path $EvidencePath ($n+'.macro.txt')));$macros+=$n}
 $db=$app.CurrentDb();$queries=@();foreach($q in $db.QueryDefs){if(-not $q.Name.StartsWith('~')){$queries+=@{name=$q.Name;sql=$q.SQL}}}
 $tables=@();foreach($t in $db.TableDefs){if(-not $t.Name.StartsWith('MSys')){$r=$db.OpenRecordset('SELECT Count(*) AS N FROM ['+$t.Name+']');$tables+=@{name=$t.Name;rows=[int]$r.Fields.Item('N').Value};$r.Close()}}
 @{source=$DatabasePath;sha256=$before;forms=$forms;modules=$modules;macros=$macros;queries=$queries;tables=$tables}|ConvertTo-Json -Depth 14|Set-Content (Join-Path $EvidencePath 'inventory.json') -Encoding UTF8
 $app.CloseCurrentDatabase()
 if((Get-FileHash -LiteralPath $DatabasePath -Algorithm SHA256).Hash -ne $before){throw 'Source modified'}
 Write-Output ('Inspected '+$forms.Count+' forms, '+$modules.Count+' modules; source unchanged: '+$DatabasePath)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
