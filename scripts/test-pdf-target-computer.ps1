param(
    [Parameter(Mandatory=$true)][string]$BuildPath,
    [string]$FixtureDirectory,
    [string]$ReportPath
)
# Run under the target user's normal account. No installs, trust changes or writes to the supplied app/PDFs.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $FixtureDirectory){$FixtureDirectory=Join-Path $root 'output/pdf/r126'}
if(-not $ReportPath){$ReportPath=Join-Path $root ('output/pdf/dependency-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.json')}
$BuildPath=[IO.Path]::GetFullPath($BuildPath);$FixtureDirectory=[IO.Path]::GetFullPath($FixtureDirectory)
$ReportPath=[IO.Path]::GetFullPath($ReportPath)
if(Test-Path -LiteralPath $ReportPath){throw "Refusing to overwrite report: $ReportPath"}
$before=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $env:TEMP ('swift-pdf-dependency-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$evidence=[ordered]@{Timestamp=(Get-Date -Format o);Computer=$env:COMPUTERNAME;Windows=[Environment]::OSVersion.VersionString;AppSHA256=$before;Access=$null;Tests=@();Status='FAIL';Error=$null}
$app=$null
try{
    $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'PDFDependencyHarness',(Join-Path $root 'tests/PDFDependencyHarness.bas'))
    $app.DoCmd.RunCommand(126);$evidence.Access=$app.Eval('PDFTargetContext()')
    foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){
        $path=Join-Path $FixtureDirectory ($kind.ToLower()+'-expanded.pdf')
        $hash=(Get-FileHash -LiteralPath $path).Hash
        $result=$app.Eval('PDFTargetProbe("'+$path.Replace('"','""')+'","'+$kind+'")')
        $evidence.Tests+=@{Kind=$kind;FixtureSHA256=$hash;Result=$result;SourceUnchanged=((Get-FileHash -LiteralPath $path).Hash -eq $hash)}
    }
    if(@($evidence.Tests|Where-Object {$_.Result -notlike 'PASS:*' -or -not $_.SourceUnchanged}).Count -eq 0){$evidence.Status='PASS'}
}catch{$evidence.Error=$_.Exception.Message}
finally{
    if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    $evidence.AppUnchanged=((Get-FileHash -LiteralPath $BuildPath).Hash -eq $before)
    if(-not $evidence.AppUnchanged){$evidence.Status='FAIL'}
    [void](New-Item -ItemType Directory -Force -Path (Split-Path $ReportPath -Parent))
    $evidence|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $ReportPath -Encoding UTF8
}
Write-Output ($evidence|ConvertTo-Json -Depth 6)
if($evidence.Status -ne 'PASS'){exit 1}
