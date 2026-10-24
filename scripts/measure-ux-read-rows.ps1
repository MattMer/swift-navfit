param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateRange(100,1000)][int]$Size=100)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/read-scale-'+[guid]::NewGuid()+'.accdb');$hash=(Get-FileHash -LiteralPath $BuildPath).Hash;Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@()
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'UXReadRowsScaleHarness',(Join-Path $root 'tests/UXReadRowsScaleHarness.bas'));$app.DoCmd.RunCommand(126)
 Write-Output ('Controller: read-row scale runner; disposable database: '+$scratch)
 $result=$app.Eval('TestReadRowsScale('+$Size+')');if($result -notlike 'PASS:*'){throw $result}
 $result.Substring(7)|Set-Content -LiteralPath ($scratch+'.timings.csv') -Encoding utf8
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){throw 'Supplied scale-test artifact changed'}
 @{Status='PASS';BuildPath=$BuildPath;BuildSHA256=$hash;MaximumRows=$Size;Trials=5;IdenticalValuesAndOrder=$true;NoReadHistoryWrites=$true;InstalledApplicationSource=$false;Timings=($scratch+'.timings.csv');Scope='Prepared versus unprepared reads on one synthetic workspace; development timings, not a release or performance guarantee'}|ConvertTo-Json|Set-Content -LiteralPath ($scratch+'.verification.json')
 Write-Output $result;Write-Output ('Receipt: '+$scratch+'.verification.json')
}catch{$primary=$_;throw}finally{
 try{$app.Quit(2)}catch{$cleanup+='Owned scale session close: '+$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+='Owned scale COM release: '+$_.Exception.Message}
 if($cleanup.Count){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
