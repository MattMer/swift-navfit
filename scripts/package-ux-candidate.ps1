param([Parameter(Mandatory=$true)][string]$TemplatePath,
[Parameter(Mandatory=$true)][string]$DemoPath,
[Parameter(Mandatory=$true)][string]$OutputDirectory,
[switch]$NoArchive,
[string]$ExpectedVersion='0.24.0-preview',
[int]$ExpectedSchema=23,
[string]$ResumeStage)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$TemplatePath=[IO.Path]::GetFullPath($TemplatePath);$DemoPath=[IO.Path]::GetFullPath($DemoPath);$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Candidate directory exists; refusing to overwrite.'}
if(-not $NoArchive -and (Test-Path -LiteralPath ($OutputDirectory+'.zip'))){throw 'Candidate archive exists; refusing to overwrite.'}
$before=@{Template=(Get-FileHash -LiteralPath $TemplatePath).Hash;Demo=(Get-FileHash -LiteralPath $DemoPath).Hash}
if($ResumeStage){
$stage=[IO.Path]::GetFullPath($ResumeStage)
if((Split-Path $stage -Parent) -ne (Split-Path $OutputDirectory -Parent) -or (Split-Path $stage -Leaf) -notlike 'candidate-stage-*'){throw 'Resume stage must be an existing candidate in the output parent'}
}else{$stage=Join-Path (Split-Path $OutputDirectory -Parent) ('candidate-stage-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $stage)}
$blank=Join-Path $stage ('Swift_NAVFIT_'+$ExpectedVersion+'-blank.accdb');$demo=Join-Path $stage ('Swift_NAVFIT_'+$ExpectedVersion+'-DEMO.accdb')
if($ResumeStage){if((Get-FileHash $blank).Hash -ne $before.Template -or (Get-FileHash $demo).Hash -ne $before.Demo){throw 'Resume package/source bytes differ'}}else{Copy-Item -LiteralPath $TemplatePath -Destination $blank;Copy-Item -LiteralPath $DemoPath -Destination $demo}
$progressPath=Join-Path $stage 'verification-progress.json';$completed=@{}
if(Test-Path $progressPath){$prior=Get-Content $progressPath -Raw|ConvertFrom-Json;if($prior.Template -ne $before.Template -or $prior.Demo -ne $before.Demo){throw 'Verification progress belongs to different artifacts'};foreach($record in $prior.Tests){$completed[$record.Name]=$record}}
$dao=New-Object -ComObject DAO.DBEngine.120;$db=$dao.OpenDatabase($blank,$false,$true);$counts=[ordered]@{}
try{
foreach($table in $db.TableDefs){if($table.Name -notlike 'MSys*' -and $table.Name -ne 'AppMetadata'){$r=$db.OpenRecordset('SELECT Count(*) FROM ['+$table.Name+']',4);$counts[$table.Name]=$r.Fields.Item(0).Value;$r.Close();if($counts[$table.Name] -ne 0){throw ('Blank candidate contains user data in '+$table.Name)}}}
$r=$db.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4);$version=$r.Fields.Item('AppVersion').Value;$schema=$r.Fields.Item('SchemaVersion').Value;$r.Close()
if($schema -ne $ExpectedSchema -or $version -ne $ExpectedVersion){throw 'Candidate metadata differs from expected preview'}
}finally{$db.Close()}
[void](New-Item -ItemType Directory -Force -Path (Join-Path $stage 'verification'))
# Validate exact packaged bytes using disposable copies; never run fixtures in the deliverable.
$checks=[ordered]@{}
foreach($test in @('test-ux-review-return.ps1','test-ux-workflows.ps1','test-ux-import-flow.ps1','test-pdf-import.ps1','test-import-review.ps1','test-import-commit.ps1','test-ux-help.ps1','test-ux-context.ps1','test-ux-shell.ps1','smoke-test.ps1','test-backup.ps1','test-major.ps1','test-preview-launch.ps1','test-ux-validation.ps1','test-ux-navywriter.ps1','test-ux-native-groups.ps1','test-directory-batch.ps1','test-directory-history.ps1','test-report-notes.ps1','test-ux-note-entries.ps1','test-ux-header-review.ps1','test-ux-global-defaults.ps1','test-ux-organization.ps1','test-ux-prior-grid.ps1','test-service-details.ps1','test-ux-responsive-identity.ps1','test-group-correctness.ps1')){
$log=Join-Path $stage ('verification/'+$test+'.log')
$testHash=(Get-FileHash (Join-Path $PSScriptRoot $test)).Hash
if($completed.ContainsKey($test) -and $completed[$test].ScriptSHA256 -eq $testHash -and (Test-Path $log) -and $completed[$test].LogSHA256 -eq (Get-FileHash $log).Hash){$checks[$test]='PASS';Write-Output ('PASS: preserved exact-byte verification '+$test);continue}

& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot $test) -BuildPath $blank *> $log
if($LASTEXITCODE -ne 0){Get-Content $log -Tail 12;throw ('Candidate verification failed: '+$test+'; retained at '+$stage)}
$checks[$test]='PASS';Write-Output ('PASS: candidate '+$test)
$completed[$test]=[ordered]@{Name=$test;ScriptSHA256=$testHash;LogSHA256=(Get-FileHash $log).Hash}
[ordered]@{Template=$before.Template;Demo=$before.Demo;Tests=@($completed.Values)}|ConvertTo-Json -Depth 5|Set-Content $progressPath -Encoding utf8
}
$log=Join-Path $stage 'verification/workspace-upgrade.log'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'test-workspace-upgrade.ps1') -TemplatePath $blank -Schema21Path $demo *> $log
if($LASTEXITCODE -ne 0){Get-Content $log -Tail 12;throw ('Candidate migration failed; retained at '+$stage)};$checks['schema20 to current / current data transfer']='PASS'
$log=Join-Path $stage 'verification/demo.log'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'test-ux-demo.ps1') -BuildPath $demo *> $log
if($LASTEXITCODE -ne 0){throw ('Demo verification failed; retained at '+$stage)};$checks['seeded demo']='PASS'
# Audit installed standard modules from a disposable candidate copy against current source.
$audit=Join-Path $root ('build/tests/candidate-audit-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $blank -Destination $audit
function NormalizeVBA([string]$text){
$text=($text.Replace("`r`n","`n") -replace '(?m)^Attribute VB_Name = .*\n','').Trim()
$parts=[regex]::Matches($text,'"(?:[^"]|"")*"|[^"]+')
$result=($parts|ForEach-Object{if($_.Value.StartsWith('"')){$_.Value}else{
$code=$_.Value.ToLowerInvariant()
$code=$code -replace '(?m)(call [a-z0-9_]+)\(\)(?=\s*(?:$|:))','$1'
$code -replace '(?m)^[ \t]+','' -replace ':[ \t]+',':'
}}) -join ''
return $result.Trim()
}
$app=New-Object -ComObject Access.Application;$moduleCount=0
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($audit,$true)
foreach($file in Get-ChildItem (Join-Path $root 'src') -Filter '*.bas'){
$export=Join-Path $stage ('verification/'+$file.BaseName+'.bas');$app.SaveAsText(5,$file.BaseName,$export)
if((NormalizeVBA ([IO.File]::ReadAllText($export))) -cne (NormalizeVBA ([IO.File]::ReadAllText($file.FullName)))){throw ('Installed code differs from current source: '+$file.BaseName)};$moduleCount++
}
foreach($form in @('frmEditor','frmUXComparison')){
$app.DoCmd.OpenForm($form,1);$f=$app.Forms.Item($form)
if($f.OnUnload -ne '[Event Procedure]'){throw ($form+' unload handler missing')}
$code=$f.Module.Lines(1,$f.Module.CountOfLines)
if($form -eq 'frmEditor' -and $code -notmatch 'If Not UXEditorUnload\(\) Then Cancel = True'){throw 'Editor save guard missing'}
if($form -eq 'frmUXComparison' -and $code -notmatch 'UXGridUnload Me'){throw 'Comparison cleanup guard missing'}
$app.DoCmd.Close(2,$form,2)
}
$app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
foreach($dir in @('scripts','src','tests','guides')){[void](New-Item -ItemType Directory -Force -Path (Join-Path $stage $dir))}
foreach($relative in @('scripts/upgrade-workspace.ps1','scripts/workspace-backup.ps1','tests/modMigration.bas','src/major-schema.json')){Copy-Item -LiteralPath (Join-Path $root $relative) -Destination (Join-Path $stage $relative)}
foreach($doc in Get-ChildItem (Join-Path $root 'planning') -Filter '*.md'){Copy-Item -LiteralPath $doc.FullName -Destination (Join-Path $stage ('guides/'+$doc.Name))}
Copy-Item -LiteralPath (Join-Path $root 'planning/UX_CANDIDATE_GUIDE.md') -Destination (Join-Path $stage 'START_HERE.md')
$packageReadme=(Get-Content (Join-Path $root 'planning/UX_CANDIDATE_GUIDE.md') -Raw)+"`r`n## Current features`r`n`r`nSee guides/UX_WORKSPACE_GUIDE.md for the complete feature-destination table.`r`n"
$packageReadme|Set-Content (Join-Path $stage 'README.md') -Encoding utf8
if((Get-FileHash -LiteralPath $blank).Hash -ne $before.Template -or (Get-FileHash -LiteralPath $demo).Hash -ne $before.Demo -or (Get-FileHash -LiteralPath $TemplatePath).Hash -ne $before.Template -or (Get-FileHash -LiteralPath $DemoPath).Hash -ne $before.Demo){throw 'Verification changed a deliverable/source'}
$files=@(Get-ChildItem -LiteralPath $stage -File -Recurse|ForEach-Object{[ordered]@{Path=$_.FullName.Substring($stage.Length+1).Replace('\','/');SHA256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
$manifest=[ordered]@{Status='Verified deployment prerelease';AppVersion=$version;SchemaVersion=$schema;BuiltAtUTC=[DateTime]::UtcNow.ToString('o');BlankUserDataTables=$counts;Checks=$checks;SourceModulesMatched=$moduleCount;SourceAndDeliverablesUnchanged=$true;Files=$files;PendingAcceptance=@('Physical 1366x768/1920x1080 displays at 100/125/150 percent scaling','Manual keyboard and screen-reader acceptance','Full NAVFIT interoperability/printed layout outside candidate verification')}
$manifest|ConvertTo-Json -Depth 8|Set-Content (Join-Path $stage 'MANIFEST.json') -Encoding utf8
if([IO.Path]::GetFullPath((Split-Path $stage -Parent)) -ne [IO.Path]::GetFullPath((Split-Path $OutputDirectory -Parent))){throw 'Staging directory escaped candidate parent'}
Move-Item -LiteralPath $stage -Destination $OutputDirectory
if(-not $NoArchive){
Compress-Archive -LiteralPath $OutputDirectory -DestinationPath ($OutputDirectory+'.zip')
(Get-FileHash -LiteralPath ($OutputDirectory+'.zip')).Hash|Set-Content -LiteralPath ($OutputDirectory+'.zip.sha256') -Encoding ascii
}
Write-Output ('PASS: local review candidate packaged at '+$OutputDirectory)
