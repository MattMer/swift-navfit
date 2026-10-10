param([Parameter(Mandatory=$true)][string]$BlankPath,[switch]$ResumeAudit)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$package=Join-Path $root 'output/Swift_NAVFIT_Baseline_2026-10-10-r98'
$blank=Join-Path $package 'Swift_NAVFIT_Deployment_2026-10-10-r98-blank.accdb'
$demo=Join-Path $package 'Swift_NAVFIT_Deployment_2026-10-10-r98-DEMO.accdb'
if((Test-Path -LiteralPath $package) -and -not $ResumeAudit){throw 'Refusing to overwrite an established baseline'}
if(-not $ResumeAudit){
[void](New-Item -ItemType Directory -Path $package)
$blank=Join-Path $package 'Swift_NAVFIT_Deployment_2026-10-10-r98-blank.accdb'
$demo=Join-Path $package 'Swift_NAVFIT_Deployment_2026-10-10-r98-DEMO.accdb'
Copy-Item -LiteralPath $BlankPath -Destination $blank
Copy-Item -LiteralPath $blank -Destination $demo
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($demo,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 $app.LoadFromText(5,'UXDemoSeed',(Join-Path $root 'tests/UXDemoSeed.bas'))
 $result=$app.Eval('SeedUXDemo()');if($result -notlike 'PASS:*'){throw $result}
 $app.DoCmd.DeleteObject(5,'UXDemoSeed')
 $app.LoadFromText(5,'UXRUADSeed',(Join-Path $root 'tests/UXRUADSeed.bas'))
 if($app.Eval('SeedRUADDemo()') -ne 2){throw 'Synthetic RUAD profile seed failed'}
 $app.DoCmd.DeleteObject(5,'UXRUADSeed');$app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
foreach($folder in @('src','scripts','tests')){Copy-Item -LiteralPath (Join-Path $root $folder) -Destination $package -Recurse}
[void](New-Item -ItemType Directory -Path (Join-Path $package 'guides'))
foreach($name in @('BASELINE_R98_2026-10-10.md','REMAINING_WORK_2026-10-09.md','RUAD_IMPORT_R97_2026-10-09.md','REVISION_CURSOR_R96_2026-10-09.md','UX_WORKSPACE_GUIDE.md')){Copy-Item -LiteralPath (Join-Path $root ('planning/'+$name)) -Destination (Join-Path $package 'guides')}
@'
# Swift NAVFIT baseline r98 — 10 October 2026

Open Swift_NAVFIT_Deployment_2026-10-10-r98-DEMO.accdb to explore synthetic reports, groups, review work, recovery and RUAD profiles. Open the matching blank file for a new workspace; it starts at command/contributor setup.

Keep each workspace in its own local folder. The file contains both application and data; separate copies do not synchronize. Preserve a copy before transferring an existing workspace. Do not replace your existing workspace with this blank file. Transfer guidance is in guides/UX_WORKSPACE_GUIDE.md; schema-20-to-25 deployment migration has not yet been accepted.

This is a development baseline (application 0.25.0-dev, schema 25), not production or Navy certification. See guides/BASELINE_R98_2026-10-10.md for tested scope and limitations. Narrative fit, full native returned-group reconciliation, target-computer PDF extraction dependencies and physical accessibility remain unfinished.

Requires Microsoft Access and compatible ACE runtime. If Access blocks active content, use your organization's approved trusted-file/location process after verifying the package. This package changes no system-wide trust settings. Native NAVFIT98A, NAVPERS templates and third-party extraction runtimes are not bundled.
'@ | Set-Content -LiteralPath (Join-Path $package 'START_HERE.md') -Encoding UTF8
}
$audit=@();$app=New-Object -ComObject Access.Application
function Normalize-CodeWhitespace([string]$value){
 $parts=[regex]::Split($value,'("(?:[^"]|"")*")')
 $result=foreach($part in $parts){if($part.StartsWith('"')){$part}else{[regex]::Replace($part,'[ \t]+',' ')}}
 [regex]::Replace(($result -join ''),'(?m)^ +| +$','').Trim()
}
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($blank,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 if($app.Eval('DCount("*","Workspace")+DCount("*","Report")+DCount("*","SummaryGroup")+DCount("*","CommandMember")') -ne 0){throw 'Blank deployment contains workspace data'}
 foreach($file in Get-ChildItem (Join-Path $root 'src') -Filter '*.bas'){
  $export=Join-Path $root ('build/tests/baseline-source-'+$file.Name)
  $app.SaveAsText(5,$file.BaseName,$export)
  $source=[regex]::Replace((Get-Content -LiteralPath $file.FullName -Raw),'(?m)^Attribute .*\r?\n','').Replace("`r`n","`n").Trim()
  $installed=[regex]::Replace((Get-Content -LiteralPath $export -Raw),'(?m)^Attribute .*\r?\n','').Replace("`r`n","`n").Trim()
  $source=[regex]::Replace($source,'(?i)\bCall\s+(\w+)\(\)','Call $1')
  $installed=[regex]::Replace($installed,'(?i)\bCall\s+(\w+)\(\)','Call $1')
  $source=Normalize-CodeWhitespace $source;$installed=Normalize-CodeWhitespace $installed
  # The VBE canonicalizes identifier case (e.g. .Value to .value) when compiling.
  if($source -ine $installed){throw ('Installed module differs from source: '+$file.BaseName)}
  $audit+=$file.BaseName
 }
 $app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
@{ModulesMatched=$audit.Count;Modules=$audit;Comparison='Attribute lines, line endings, VBE case, code whitespace outside strings and empty Call parentheses normalized';BlankWorkspaceRows=0;AppVersion='0.25.0-dev';SchemaVersion=25}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $package 'source-audit.json') -Encoding UTF8
Write-Output ('BASELINE_PACKAGE_READY: '+$package)
