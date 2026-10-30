param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$dir=Join-Path $root ('build/tests/navfit-adapter-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$original=Join-Path $root 'FORMS_REFERENCES_TEMPLATES/TEST_DB.accdb';$hash=(Get-FileHash -LiteralPath $original).Hash
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path (Split-Path $PSScriptRoot -Parent) 'make-navfit-full-fixture.ps1') -DestinationPath $source
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@()
try{
 Write-Output ('Controller: NAVFIT adapter runner; disposable database: '+$scratch)
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallSource){foreach($m in @('modNAVFITMapping','modNAVFITAdapter','modNAVFITExchange','modImportReview')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas'))) }}
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
 Write-Output $result
}catch{$primary=$_}finally{
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $buildHash){$cleanup+='Source build changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
@{Status=$result;ExchangeStatus=$exchange;Guards=$guards;SourceUnchanged=$true;BuildSHA256=$buildHash;InstalledSource=[bool]$InstallSource;NativeApplicationProof=$false;Directory=$dir}|ConvertTo-Json|Set-Content (Join-Path $dir 'verification.json')
