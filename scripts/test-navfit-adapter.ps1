param([string]$BuildPath='D:\Development\Swift NAVFIT\build\ux-r26-final.accdb')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dir=Join-Path $root ('build/tests/navfit-adapter-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$original=Join-Path $root 'FORMS_REFERENCES_TEMPLATES/TEST_DB.accdb';$hash=(Get-FileHash -LiteralPath $original).Hash
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $source
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($m in @('modNAVFITMapping','modNAVFITAdapter','modNAVFITExchange','modImportReview')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
 $app.LoadFromText(5,'NAVFITAdapterHarness',(Join-Path $root 'tests/NAVFITAdapterHarness.bas'))
 $app.LoadFromText(5,'NAVFITExchangeHarness',(Join-Path $root 'tests/NAVFITExchangeHarness.bas'))
 $app.LoadFromText(5,'NAVFITGuardHarness',(Join-Path $root 'tests/NAVFITGuardHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestNAVFITAdapter("'+$source+'","'+(Join-Path $dir 'roundtrip.accdb')+'")')
 if($result -notlike 'PASS:*'){throw $result}
 $exchange=$app.Eval('TestNAVFITExchange("'+$dir+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")')
 if($exchange -notlike 'PASS:*'){throw $exchange}
 Write-Output $exchange
 $guards=$app.Eval('TestNAVFITGuards("'+$dir+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")')
 if($guards -notlike 'PASS:*'){throw $guards}
 Write-Output $guards
 $app.CloseCurrentDatabase()
 if((Get-FileHash -LiteralPath $original).Hash -ne $hash){throw 'Source changed'}
 @{Status=$result;ExchangeStatus=$exchange;Guards=$guards;SourceUnchanged=$true;NativeApplicationProof=$false;Directory=$dir}|ConvertTo-Json|Set-Content (Join-Path $dir 'verification.json')
 Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
