param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$SamplePath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if(-not $SamplePath){
 #Synthetic NRRM roster in the verified 35-column format: 60 people, then 6 vacant billets.
 $keys='CUI|RUIC|RESERVE UNIT|RESERVE ACTIVITY AND LOC|ST|UIC|NRC|NRA|RPC|RATE/RANK|NAME|RFAS Asgmt|E/O|SELRES TYPE|BIN|RFAS|REQ DESG/REQ RTING|PNOBC/PNEC|DESG|NOBC/NEC|AUIC|TRUIC|UMUIC|MOB ACTIVITY|MOB BILLET|PRD|IMS|MAS (A)|MAS (M)|MAS (T)|ABSC|RBSC|ADDRESS|PHONE|NROWS Email'.Split('|')
 $lines=@($keys -join ',')
 foreach($i in 1..66){
  $row=[ordered]@{};foreach($k in $keys){$row[$k]='SYN'}
  $row['UIC']='0'+(12340+$i%3);$row['TRUIC']=$row['UIC'];$row['UMUIC']=$row['UIC'];$row['PRD']='2027-0'+(1+$i%9)
  $row['PHONE']='[PII]';$row['ADDRESS']='[PII]';$row['NROWS Email']='[PII]'
  if($i -le 60){$row['NAME']='SYNTHETIC MEMBER '+$i.ToString('00');$row['RATE/RANK']='IT2';$row['DESG']='IT2';$row['E/O']='Enlisted';$row['SELRES TYPE']='Drilling'}
  else{$row['NAME']='';$row['RATE/RANK']='';$row['DESG']='';$row['E/O']='Enlisted';$row['SELRES TYPE']='Vacant'}
  $lines+=($row.Values -join ',')
 }
 $SamplePath=Join-Path $root ('build/tests/ruad-sample-'+[guid]::NewGuid()+'.csv')
 [IO.File]::WriteAllText($SamplePath,($lines -join "`r`n")+"`r`n",(New-Object Text.UTF8Encoding($false)))
}
$scratch=Join-Path $root ('build/tests/ruad-import-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 $app.LoadFromText(5,'RUADImportHarness',(Join-Path $root 'tests/RUADImportHarness.bas'));$app.DoCmd.RunCommand(126)
 $expression='TestRUADImport("'+$SamplePath.Replace('"','""')+'")'
 $result=$app.Eval($expression);$result;if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
