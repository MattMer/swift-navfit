param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$LoadCurrentSource)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/trait-retry-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@()
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 if($LoadCurrentSource){$app.LoadFromText(5,'modUXTraitGrid',(Join-Path $root 'src/modUXTraitGrid.bas'))}
 $app.LoadFromText(5,'UXTraitRetryHarness',(Join-Path $root 'tests/UXTraitRetryHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestTraitRetry()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
} catch {$primary=$_} finally {
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){$cleanup+='Trait retry changed source artifact'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
