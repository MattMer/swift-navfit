param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dir=Join-Path $root ('build/tests/navfit-group-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $source
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@();$results=@()
try{
 Write-Output ('Controller: NAVFIT group runner; disposable database: '+$scratch)
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallSource){foreach($m in @('modNAVFITExchange','modUXNAVFIT')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas'))) }}
 foreach($m in @('NAVFITAdapterHarness','NAVFITGroupExportHarness')){$app.LoadFromText(5,$m,(Join-Path $root ('tests/'+$m+'.bas')))}
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestNAVFITAdapter("'+$source+'","'+(Join-Path $dir 'roundtrip.accdb')+'")');if($result -notlike 'PASS:*'){throw $result};Write-Output $result;$results+=$result
 foreach($test in @('TestNAVFITGroupExport','TestNAVFITNonblocking')){
  $result=$app.Eval($test+'("'+$dir+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")');if($result -notlike 'PASS:*'){throw $result};Write-Output $result;$results+=$result
 }
}catch{$primary=$_}finally{
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){$cleanup+='Source build changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
@{Results=$results;BuildSHA256=$hash;SourceUnchanged=$true;InstalledSource=[bool]$InstallSource;Directory=$dir}|ConvertTo-Json|Set-Content (Join-Path $dir 'verification.json')
