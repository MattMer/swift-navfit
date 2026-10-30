param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$dir=Join-Path $root ('build/tests/navfit-ui-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path (Split-Path $PSScriptRoot -Parent) 'make-navfit-full-fixture.ps1') -DestinationPath $source
$test=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $test
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@()
try{
 Write-Output ('Controller: NAVFIT UI runner; disposable database: '+$test)
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($test,$true)
 if($InstallSource){foreach($m in @('modNAVFITExchange','modUXNAVFIT','modUXExportNames')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas'))) }}
 $app.LoadFromText(5,'UXNAVFITHarness',(Join-Path $root 'tests/UXNAVFITHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestUXNAVFIT("'+$source+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")');if($result -notlike 'PASS:*'){throw $result}
 $app.CloseCurrentDatabase()
 Write-Output $result
}catch{$primary=$_}finally{
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){$cleanup+='Source build changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
@{Status=$result;TestPath=$test;BuildSHA256=$hash;SourceUnchanged=$true;InstalledSource=[bool]$InstallSource}|ConvertTo-Json|Set-Content (Join-Path $dir 'verification.json')
