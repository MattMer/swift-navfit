param(
    [string]$OutputPath=(Join-Path $PSScriptRoot ('../output/ux/Swift_NAVFIT_UI_Demo_'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.accdb')),
    [string]$TemplatePath
)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw "Refusing to overwrite: $OutputPath"}
if(-not $TemplatePath){
    $TemplatePath=Join-Path $root ('build/ux-blank-'+[guid]::NewGuid()+'.accdb')
    & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'build.ps1') -OutputPath $TemplatePath
    if($LASTEXITCODE -ne 0){throw 'Native template build failed'}
}
$sourceHash=(Get-FileHash -LiteralPath $TemplatePath).Hash
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'test-ux-surface.ps1') -BuildPath $TemplatePath
if($LASTEXITCODE -ne 0){throw 'Native editor/workspace verification failed; demo not created'}
[void](New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath -Parent))
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
    $app.LoadFromText(5,'UXDemoSeed',(Join-Path $root 'tests/UXDemoSeed.bas'))
    $result=$app.Eval('SeedUXDemo()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.DoCmd.DeleteObject(5,'UXDemoSeed')
    $app.CloseCurrentDatabase()
    $app.OpenCurrentDatabase($OutputPath,$true)
    if(-not $app.Eval('DCount("*","Report")=9 And DCount("*","SummaryGroup")=3 And DCount("*","RecoveryDraft")=1')){throw 'Demo counts changed on reopen'}
    $app.CloseCurrentDatabase()
}finally{
    if($null -ne $app){
        $app.Quit(2)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        $app=$null
    }
}
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'test-ux-demo.ps1') -BuildPath $OutputPath
if($LASTEXITCODE -ne 0){throw 'Populated preview verification failed'}
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'test-preview-launch.ps1') -BuildPath $OutputPath
if($LASTEXITCODE -ne 0){throw 'Preview startup verification failed'}
[pscustomobject]@{Output=$OutputPath;SHA256=(Get-FileHash -LiteralPath $OutputPath).Hash;Status=$result;SourceUnchanged=((Get-FileHash -LiteralPath $TemplatePath).Hash -eq $sourceHash)}|ConvertTo-Json
