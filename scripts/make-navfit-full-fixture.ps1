param([Parameter(Mandatory=$true)][string]$DestinationPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$source=Join-Path $root 'FORMS_REFERENCES_TEMPLATES/TEST_DB.accdb';$hash=(Get-FileHash -LiteralPath $source).Hash
if(Test-Path -LiteralPath $DestinationPath){throw 'Refusing to overwrite fixture'}
Copy-Item -LiteralPath $source -Destination $DestinationPath
$map=Get-Content -Raw (Join-Path $root 'src/navfit-field-map.json')|ConvertFrom-Json
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($DestinationPath)
try{
 $r=$db.OpenRecordset('SELECT * FROM Reports WHERE ReportType Is Not Null ORDER BY ReportID',2)
 while(-not $r.EOF){
  $kind=switch($r.Fields.Item('ReportType').Value){'Eval'{'EVAL'}'Chief'{'CHIEFEVAL'}'FitRep'{'FITREP'}}
  $assignments=@()
  foreach($row in $map.$kind){
   $f=$r.Fields.Item($row.source)
   $value=switch($row.conversion){
    'date'{[datetime]'2026-09-30'}
    'counseling'{'26JAN01'}
    'flag'{[bool]($row.target -in @('ComponentActive','OccasionPeriodic','ReportRegular','RetentionYes','StatementNo'))}
    'observation'{$false}
    'trait'{[int](3+([int]$row.target.Substring(5)%3))}
    'promotion'{4}
    'text'{('SYNTHETIC '+$row.source).Substring(0,[Math]::Min($f.Size,('SYNTHETIC '+$row.source).Length))}
   }
   if($f.Type -eq 12){$value='SYNTHETIC '+$kind+' '+$row.source+"`r`nSecond line; apostrophe ' and & retained."}
   if($row.target -match 'SSN$'){$value='00000000'+$r.Fields.Item('ReportID').Value}
   if($row.target -match 'DoDID$'){$value='0000000001'}
   if($row.target -match 'UIC$'){$value='12345'}
   if($row.target -eq 'PromotionStatus'){$value='REGULAR'}
   if($row.target -eq 'RankRate'){$value=switch($kind){'EVAL'{'IT2'}'CHIEFEVAL'{'ITC'}'FITREP'{'LT'}}}
   if($row.target -eq 'ReportingSeniorGrade'){$value='CAPT'}
   if($row.target -eq 'ReportingSeniorState'){$value='VA'}
   if($row.target -eq 'PeriodStart'){$value=[datetime]'2026-01-01'}
   if($value -is [datetime]){$literal='#'+$value.ToString('yyyy-MM-dd')+'#'}elseif($value -is [bool]){$literal=if($value){'True'}else{'False'}}elseif($value -is [int]){$literal=[string]$value}else{$literal="'"+([string]$value).Replace("'","''")+"'"};$assignments+=('['+$row.source+']='+$literal)
  }
  $r.Edit()
  $r.Fields.Item('Achievements').Value='SYNTHETIC COMMAND ACHIEVEMENTS '+$kind
  $r.Fields.Item('Qualifications').Value='SYNTHETIC QUALIFICATIONS '+$kind
  $r.Fields.Item('RSAddress').Value="SYNTHETIC combined address`r`nPreserve original"
  $r.Fields.Item('RSInfo').Value='SYNTHETIC concurrent reporting senior legacy text'
  $r.Fields.Item('UserComments').Value='SYNTHETIC native local notes'
  $r.Fields.Item('PROFDN1').Value=''
  $r.Fields.Item('UserComments').Value='Unicode '+[char]0x03A9+' '+[char]0x00E9+' '+('long native memo; ' * 400)
  $r.Fields.Item('IsValidated').Value=$false
  $r.Update();$db.Execute('UPDATE Reports SET '+($assignments -join ',')+' WHERE ReportID='+$r.Fields.Item('ReportID').Value,128);$r.MoveNext()
 }
 $r.Close()
}finally{$db.Close()}
if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Original changed'}
Write-Output $DestinationPath
