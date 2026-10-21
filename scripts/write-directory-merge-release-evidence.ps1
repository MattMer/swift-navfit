param([Parameter(Mandatory=$true)][string]$BlankPath,[Parameter(Mandatory=$true)][string]$DemoPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$native=Get-Content (Join-Path $root 'build/directory-merge-release-verification.json') -Raw|ConvertFrom-Json
if($native.Status -ne 'PASS'){throw 'Native release gates incomplete'}
$files=@();$dao=New-Object -ComObject DAO.DBEngine.120
try{
 foreach($pair in @(@('Blank',$BlankPath),@('Demo',$DemoPath))){
  $path=[IO.Path]::GetFullPath($pair[1]);$db=$dao.OpenDatabase($path,$false,$true)
  try{
   $r=$db.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4);$version=$r.Fields.Item(0).Value;$schema=[int]$r.Fields.Item(1).Value;$r.Close()
   if($version -ne '0.27.0-dev' -or $schema -ne 27){throw 'Wrong release identity'}
   $counts=[ordered]@{}
   foreach($name in @($db.TableDefs|ForEach-Object {$_.Name}|Where-Object {$_ -notlike 'MSys*'}|Sort-Object)){
    $r=$db.OpenRecordset(('SELECT Count(*) FROM ['+$name+']'),4);$counts[$name]=[int]$r.Fields.Item(0).Value;$r.Close()
   }
   if($pair[0] -eq 'Blank'){foreach($name in $counts.Keys){if($name -ne 'AppMetadata' -and $counts[$name] -ne 0){throw ('Blank contains data: '+$name)}}}
   if($pair[0] -eq 'Demo' -and ($counts.Report -ne 13 -or $counts.SummaryGroup -ne 5 -or $counts.CommandMember -ne 13)){throw 'Unexpected Demo data'}
   foreach($name in @('DirectoryMerge','DirectoryMergeReport','DirectoryMergePrior')){if($counts[$name] -ne 0){throw 'Scratch merge data leaked into distribution'}}
  }finally{$db.Close()}
  $files+=@{kind=$pair[0];path=$path;name=[IO.Path]::GetFileName($path);bytes=(Get-Item -LiteralPath $path).Length;sha256=(Get-FileHash -LiteralPath $path).Hash;appVersion=$version;schema=$schema;counts=$counts}
 }
}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao)}
if($files[0].sha256 -ne $native.BlankSHA256 -or $files[1].sha256 -ne $native.DemoSHA256){throw 'Artifact changed after native gates'}
$originals=@()
foreach($entry in @(
 @('output/Swift_NAVFIT_Blank_2026-10-20-r126.accdb','B5C3767436DB739A098CFE2A093F1EABF514FCEA49D4627879C7808769300992'),
 @('output/Swift_NAVFIT_Demo_2026-10-20-r126.accdb','882E50D7B2757E2497F6E54A27A462DB10CBDB9C241A9E4172D026FA3AB80623'),
 @('output/ux/Swift_NAVFIT_UI_Blank_2026-10-17-r116.accdb','6ACE66A9C2883E1BE6CBCFBEC584249B254A61620D998486468E205BCECBEFE2')
)){
 $path=Join-Path $root $entry[0];$hash=(Get-FileHash -LiteralPath $path).Hash
 $originals+=@{Path=$path;ExpectedSHA256=$entry[1];SHA256=$hash;MatchesRecordedBaseline=($hash -eq $entry[1])}
 if($hash -ne $entry[1]){
  $migration=Get-Content (Join-Path $root 'build/r115-data-import-verification.json') -Raw|ConvertFrom-Json
  $retained=@($migration.Fixtures|Where-Object {$_.Source -eq $path -and $_.SourceSHA256 -eq $hash -and $_.Result -like 'PASS:*'})
  if($retained.Count -ne 1){throw ('Original changed without a matching before/after preservation proof: '+$path)}
  $originals[-1].UnchangedDuringReadOnlyTransfer=$true
  $originals[-1].Note='Current original differs from older published receipt; cause not established. Preserved as found; exact native transfer and before/after hash passed.'
 }
}
$build=Get-Content ($files[0].path+'.build.json') -Raw|ConvertFrom-Json
$transfer=Get-Content (Join-Path $root 'build/directory-merge-transfer-verification.json') -Raw|ConvertFrom-Json
$allSchemas=Get-Content (Join-Path $root 'build/r115-data-import-verification.json') -Raw|ConvertFrom-Json
if(@($allSchemas.Fixtures).Count -ne 27){throw 'Expected all 27 retained schema gates'}
$checks=@();foreach($test in $native.Tests){$checks+=@{Name=$test.Script;Status=$test.Status;Log=$test.Log;Result=[IO.File]::ReadAllText($test.Log)}}
$manifest=@{date='2026-10-21';revision='r127';verification=(Join-Path $root 'planning/DIRECTORY_MERGE_R127_VERIFICATION.json');files=$files}
$receipt=@{Revision='r127';Date='2026-10-21';Status='PASS';AppVersion='0.27.0-dev';Schema=27;C01='User accepted folder permissions; no additional application encryption/password';Build=$build.Verification;BuildKind=$build.BuildKind;Files=$files;Checks=$checks;MergeTransfer=$transfer;RetainedSchemas=$allSchemas;Originals=$originals;PhysicalDisplayAcceptance='Deferred';FullFreshGeneration='Not established; verified incremental update on r126'}
$receipt|ConvertTo-Json -Depth 14|Set-Content (Join-Path $root 'planning/DIRECTORY_MERGE_R127_VERIFICATION.json')
$manifest|ConvertTo-Json -Depth 9|Set-Content (Join-Path $root 'output/CURRENT_FILES.json')
Write-Output 'PASS: final Blank/Demo identity, empty/synthetic data, no scratch merges, original hashes and native receipts'
