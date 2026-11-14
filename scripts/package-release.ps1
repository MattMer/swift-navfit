param(
    [Parameter(Mandatory=$true)][string]$BlankPath,
    [Parameter(Mandatory=$true)][string]$DemoPath,
    [Parameter(Mandatory=$true)][string]$OutputFolder
)
# Release files from a checked Blank and Demo (backlog S01, owner decisions 13 Nov 2026):
#   Swift_NAVFIT_<version>.accdb            the application; setup makes its data file
#   Swift_NAVFIT_Demo_<version>.accdb       the application, linked to the Demo data file
#   Swift_NAVFIT_Demo_data_<version>.accdb  the Demo data (synthetic only)
# The Demo data comes in through the production import, every row and field checked.
# The Demo application finds its data file beside it, so the pair opens wherever the
# two are saved together; this script checks that by opening a moved copy.
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
$plan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
$version=$plan.AppVersion
$BlankPath=[IO.Path]::GetFullPath($BlankPath);$DemoPath=[IO.Path]::GetFullPath($DemoPath);$OutputFolder=[IO.Path]::GetFullPath($OutputFolder)
$names=@{App='Swift_NAVFIT_'+$version+'.accdb';Demo='Swift_NAVFIT_Demo_'+$version+'.accdb';Data='Swift_NAVFIT_Demo_data_'+$version+'.accdb'}
[void](New-Item -ItemType Directory -Force -Path $OutputFolder)
foreach($name in $names.Values){if(Test-Path -LiteralPath (Join-Path $OutputFolder $name)){throw ('Refusing to overwrite '+$name)}}
$blankHash=(Get-FileHash -LiteralPath $BlankPath).Hash;$demoHash=(Get-FileHash -LiteralPath $DemoPath).Hash
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'verify-build-artifact.ps1') -BuildPath $BlankPath -RequireBlank
if($LASTEXITCODE -ne 0){throw 'The Blank is not a verified empty build'}

$app=Join-Path $OutputFolder $names.App
Copy-Item -LiteralPath $BlankPath -Destination $app
Write-Output ('Application: '+$names.App)

$demo=Join-Path $OutputFolder $names.Demo;$data=Join-Path $OutputFolder $names.Data
Copy-Item -LiteralPath $BlankPath -Destination $demo
$access=New-Object -ComObject Access.Application
try{
    $access.AutomationSecurity=1;$access.OpenCurrentDatabase($demo,$true)
    if(-not $access.Eval('UXUseDataFile("'+$data+'",True)')){throw 'The Demo data file could not be made'}
    $result=[string]$access.Eval('ImportSwiftWorkspaceData("'+$DemoPath+'","'+$demoHash+'")')
    if($result -notlike 'PASS:*'){throw ('Demo import: '+$result)}
    Write-Output ('Demo data imported: '+$result)
}finally{$access.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($access)}

#The pair, saved together somewhere else, opens on its own data file without asking.
$engine=New-Object -ComObject DAO.DBEngine.120
function Count([string]$Path,[string]$Table){$db=$engine.OpenDatabase($Path,$false,$true);try{$r=$db.OpenRecordset('SELECT Count(*) FROM ['+$Table+']',4);[int]$r.Fields.Item(0).Value}finally{$db.Close()}}
foreach($table in 'Report','CommandMember','SummaryGroup'){if((Count $DemoPath $table) -ne (Count $data $table)){throw ('Demo data file '+$table+' rows differ from the Demo')}}
$check=Join-Path $OutputFolder ('pair-check-'+[guid]::NewGuid().ToString('N').Substring(0,8))
[void](New-Item -ItemType Directory -Path $check)
#The data file is put aside meanwhile, so the moved copy cannot fall back on this one.
$aside=$data+'.packaging'
try{
    Copy-Item -LiteralPath $demo,$data -Destination $check
    Move-Item -LiteralPath $data -Destination $aside
    $access=New-Object -ComObject Access.Application
    try{
        $access.AutomationSecurity=1;$access.OpenCurrentDatabase((Join-Path $check $names.Demo),$true)
        $found=[string]$access.Eval('SwiftDataPath()')
        if($found -ne (Join-Path $check $names.Data)){throw ('The moved Demo did not find its data file beside it: '+$found)}
        if($access.Forms.Count -eq 0){throw 'The moved Demo did not start'}
    }finally{$access.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($access)}
}finally{
    if(Test-Path -LiteralPath $aside){Move-Item -LiteralPath $aside -Destination $data}
    Remove-Item -LiteralPath $check -Recurse -Force
}
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)
if((Get-FileHash -LiteralPath $BlankPath).Hash -ne $blankHash -or (Get-FileHash -LiteralPath $DemoPath).Hash -ne $demoHash){throw 'Packaging changed the Blank or the Demo'}
foreach($name in 'App','Demo','Data'){$file=Join-Path $OutputFolder $names[$name];Write-Output ($names[$name]+'  '+(Get-Item -LiteralPath $file).Length+' bytes  SHA256 '+(Get-FileHash -LiteralPath $file).Hash)}
Write-Output ('PASS: release files for '+$version+' in '+$OutputFolder+'; the Demo pair opens on its own data file after a move')
