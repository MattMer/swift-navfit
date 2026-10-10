param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$CandidatePath)
$ErrorActionPreference='Stop';$engine=New-Object -ComObject DAO.DBEngine.120
$left=$engine.OpenDatabase((Resolve-Path -LiteralPath $SourcePath).Path,$false,$true)
$right=$engine.OpenDatabase((Resolve-Path -LiteralPath $CandidatePath).Path,$false,$true)
try {
 $count=0
 foreach($table in $left.TableDefs){
  if($table.Name.StartsWith('MSys')){continue}
  $all=@()
  foreach($db in @($left,$right)){
   $r=$db.OpenRecordset('SELECT * FROM ['+$table.Name+']');$records=@()
   while(-not $r.EOF){$values=@();foreach($f in $r.Fields){$v=$f.Value;if($v -is [DBNull]){$v=$null};$values+=,$v};$records+=($values|ConvertTo-Json -Compress -Depth 5);$r.MoveNext()}
   $r.Close();$all+=(@($records|Sort-Object -CaseSensitive) -join "`n")
  }
  if($all[0] -cne $all[1]){throw ('Data mismatch '+$table.Name)};$count++
 }
 Write-Output ('PASS: all '+$count+' user tables identical')
}finally{$left.Close();$right.Close()}
