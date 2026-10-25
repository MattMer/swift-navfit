param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
#The finalizer re-checks the build inputs, so load the plan exactly as build.ps1 does.
. (Join-Path $PSScriptRoot 'build-plan.ps1')
$swiftBuildPlan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
$swiftBuildTrace=[Collections.Generic.List[object]]::new()
$OutputPath=Join-Path $root ('build/tests/finalization-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $OutputPath
$source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
$finalizer=$source.Substring($source.LastIndexOf("`n} finally {"))
# Exercise the exact finalizer with references that retained the database lock.
$fixture=@(
 '$app=$null'
 'try{'
 ' $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)'
 ' $db=$app.CurrentDb();$r=$db.OpenRecordset(''SELECT * FROM AppMetadata'')'
 ' $app.DoCmd.OpenForm(''frmUXComparison'',1);$f=$app.Forms.Item(''frmUXComparison'');$module=$f.Module;$script:swiftControl=$f.Controls.Item(''cmdBack'')'
 ' $app.DoCmd.Close(2,''frmUXComparison'',2);$app.CloseCurrentDatabase()'
) -join "`r`n"
Invoke-Expression ($fixture+$finalizer)
if(-not $builtHash -or $builtHash -ne (Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash){throw 'Final artifact hash not established'}
if((Get-Content -LiteralPath ($OutputPath+'.build.json') -Raw|ConvertFrom-Json).SHA256 -ne $builtHash){throw 'Build receipt does not record the final hash'}
Write-Output 'PASS: real build shutdown/hash finalizer with retained DAO, recordset, form, class and control references'
