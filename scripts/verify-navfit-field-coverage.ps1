param([string]$OutputPath='D:\Development\Swift NAVFIT\planning\evidence\navfit\field-coverage.json')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$map=Get-Content -Raw (Join-Path $root 'src/navfit-field-map.json')|ConvertFrom-Json
$il=Get-Content -Raw (Join-Path $root 'planning/evidence/navfit/native-bindings.json')|ConvertFrom-Json
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase((Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb'),$false,$true)
$result=[ordered]@{NativeVersion=$il.Version;NativeExecutableSHA256=$il.SHA256;NativeReferenceSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'NAVFIT98A/Data/NF98A_Dte.accdb')).Hash;Status='Binding and storage evidence; native journey verification is separate';Forms=@{}}
try{
 foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){
  $method=switch($kind){'EVAL'{'LoadEvaluationForm'}'CHIEFEVAL'{'LoadChiefForm'}'FITREP'{'LoadFitRepForm'}}
  $strings=@($il.Methods|Where-Object{$_.Method -eq $method}|ForEach-Object{$_.Instructions}|Where-Object{$_.Op -eq 'ldstr'}|ForEach-Object{$_.Operand.ToLowerInvariant()})
  $rows=@()
  foreach($f in $db.TableDefs.Item('Reports').Fields){
   $binding=@($map.$kind|Where-Object{$_.source -ieq $f.Name})
   $classification=if($binding.Count){'Mapped saved value'}elseif($f.Name -in @('ReportID','ReportType','Parent','IsValidated')){'Native identity or validation state'}else{'Retained native-only value'}
   $rows+=@{NativeField=$f.Name;DAOType=$f.Type;Size=$f.Size;SwiftFields=@($binding.target);Conversion=@($binding.conversion);ReferencedByNativeLoader=($f.Name.ToLowerInvariant() -in $strings);Classification=$classification}
  }
  $expected=if($kind -eq 'CHIEFEVAL'){96}else{75}
  if($rows.Count -ne 128 -or @($map.$kind).Count -ne $expected){throw ("Unexpected coverage: $kind rows=$($rows.Count) map=$(@($map.$kind).Count)")}
  $result.Forms[$kind]=$rows
 }
}finally{$db.Close()}
$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Output ('Covered 128 native columns; 75 EVAL/FITREP and 96 Chief saved mappings: '+$OutputPath)
