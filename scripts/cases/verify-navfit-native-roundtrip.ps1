param([Parameter(Mandatory=$true)][string]$BeforePath,[Parameter(Mandatory=$true)][string]$AfterPath,[Parameter(Mandatory=$true)][string]$EvidencePath)
$ErrorActionPreference='Stop'
$engine=New-Object -ComObject DAO.DBEngine.120;$before=$engine.OpenDatabase($BeforePath,$false,$true);$after=$engine.OpenDatabase($AfterPath,$false,$true)
$result=[ordered]@{BeforeSHA256=(Get-FileHash -LiteralPath $BeforePath).Hash;AfterSHA256=(Get-FileHash -LiteralPath $AfterPath).Hash;ComparedCells=0;Changes=@();Status='Cell comparison only; native UI actions must be recorded separately'}
try{
 foreach($table in @('Folders','Reports','Summary')){
  $order=if($table -eq 'Folders'){' ORDER BY FolderID'}elseif($table -eq 'Reports'){' ORDER BY ReportID'}else{''}
  $a=$before.OpenRecordset('SELECT * FROM ['+$table+']'+$order,2);$b=$after.OpenRecordset('SELECT * FROM ['+$table+']'+$order,2);$row=0
  while(-not $a.EOF){
   if($b.EOF){throw 'Row removed'}
   for($i=0;$i -lt $a.Fields.Count;$i++){
    $f=$a.Fields.Item($i);$v=$f.Value;$new=$b.Fields.Item($i).Value
    $nullBefore=$v -is [DBNull] -or $null -eq $v;$nullAfter=$new -is [DBNull] -or $null -eq $new
    $same=($nullBefore -eq $nullAfter) -and ($nullBefore -or ([string]$v -ceq [string]$new))
    $result.ComparedCells++
    if(-not $same){$result.Changes+=@{Table=$table;Row=$row;Field=$f.Name;DAOType=$f.Type;BeforeNull=$nullBefore;AfterNull=$nullAfter;Before=[string]$v;After=[string]$new}}
   }
   $a.MoveNext();$b.MoveNext();$row++
  }
  if(-not $b.EOF){throw 'Row added'}
  $a.Close();$b.Close()
 }
}finally{$before.Close();$after.Close()}
$result|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $EvidencePath -Encoding UTF8
Write-Output ('Compared '+$result.ComparedCells+' cells; native changed '+$result.Changes.Count+' values. See '+$EvidencePath)
