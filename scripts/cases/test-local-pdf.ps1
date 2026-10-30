param([Parameter(Mandatory=$true)][string]$BuildPath,[Parameter(Mandatory=$true)][string]$FixturePath,[switch]$InstallReader)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/local-pdf-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null;$passed=0
try{
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallReader){foreach($module in @('modPDFInflate','modPDFCrypto','modPDFLocal','modPDFImport')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
 $app.LoadFromText(5,'LocalPDFHarness',(Join-Path $root 'tests/LocalPDFHarness.bas'))
 foreach($case in (Get-Content (Join-Path $FixturePath 'manifest.json') -Raw|ConvertFrom-Json)){
  $path=Join-Path $FixturePath $case.file
  $before=(Get-FileHash -LiteralPath $path).Hash
  if($before -ne $case.sourceHash){throw ('Fixture hash differs: '+$case.file)}
  $result=$app.Eval('CheckLocalPDF("'+$path+'","'+$case.kind+'","'+$case.templateHash+'","'+$case.dataHash+'","'+$case.rejection+'")')
  if($result -notlike 'PASS:*'){throw ($case.file+': '+$result)}
  if((Get-FileHash -LiteralPath $path).Hash -ne $before){throw ('Source changed: '+$case.file)}
  $passed++;Write-Output ($case.file+': '+$result)
 }
 Write-Output ('PASS: '+$passed+' native local PDF cases; all source hashes unchanged')
}finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
