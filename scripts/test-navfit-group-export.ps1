param([string]$BuildPath='D:\Development\Swift NAVFIT\build\ux-r28-export2.accdb')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dir=Join-Path $root ('build/tests/navfit-group-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $source
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($m in @('modNAVFITExchange','modUXNAVFIT')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
 foreach($m in @('NAVFITAdapterHarness','NAVFITGroupExportHarness')){$app.LoadFromText(5,$m,(Join-Path $root ('tests/'+$m+'.bas')))}
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestNAVFITAdapter("'+$source+'","'+(Join-Path $dir 'roundtrip.accdb')+'")');if($result -notlike 'PASS:*'){throw $result};Write-Output $result
 $result=$app.Eval('TestNAVFITGroupExport("'+$dir+'","'+(Join-Path $root 'NAVFIT_BLANK.accdb')+'")');if($result -notlike 'PASS:*'){throw $result};Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
