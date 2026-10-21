param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
# Schema-27 directory merge update on the verified schema-26 blank UI.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
. (Join-Path $PSScriptRoot 'vba-source-normalization.ps1')
. (Join-Path $PSScriptRoot 'access-properties.ps1')
$swiftBuildPlan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
$swiftBuildPlan.InputHashes[$PSCommandPath]=(Get-FileHash -LiteralPath $PSCommandPath).Hash
$OutputPath=[IO.Path]::GetFullPath($OutputPath);$TemplatePath=[IO.Path]::GetFullPath($TemplatePath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing to overwrite an existing file'}
$sourceHash=(Get-FileHash -LiteralPath $TemplatePath).Hash
$dao=New-Object -ComObject DAO.DBEngine.120;$db=$dao.OpenDatabase($TemplatePath,$false,$true)
try{
 $row=$db.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4);$schema=[int]$row.Fields.Item(0).Value;$row.Close()
 if($schema -ne 26 -or $swiftBuildPlan.SchemaVersion -ne 27){throw 'This updater requires a schema-26 blank and schema-27 source'}
 foreach($table in @($db.TableDefs|ForEach-Object {$_.Name}|Where-Object {$_ -notlike 'MSys*' -and $_ -ne 'AppMetadata'})){
  $row=$db.OpenRecordset(('SELECT TOP 1 * FROM ['+$table+']'),4);$hasRows=-not $row.EOF;$row.Close()
  if($hasRows){throw ('Template must be blank: '+$table)}
 }
}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao)}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=$null;$updated=@();$exportFolder=Join-Path $root ('build/source-update-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory $exportFolder)
try{
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 $names=@($app.CurrentProject.AllModules|ForEach-Object {$_.Name})
 foreach($source in $swiftBuildPlan.Manifest.sourceModules){
  $name=[IO.Path]::GetFileNameWithoutExtension($source);$same=$false
  if($names -contains $name){
   $export=Join-Path $exportFolder ($name+'.bas');$app.SaveAsText(5,$name,$export)
   $same=(NormalizeSwiftVBA (Get-Content $export -Raw)) -ceq (NormalizeSwiftVBA (Get-Content (Join-Path $root $source) -Raw))
  }
  if(-not $same){$app.LoadFromText(5,$name,(Join-Path $root $source));$updated+=$name}
 }
 . (Join-Path $PSScriptRoot 'apply-directory-merge.ps1')
 $db=$app.CurrentDb();$db.Execute("UPDATE AppMetadata SET AppVersion='$($swiftBuildPlan.AppVersion)'",128);$db.Close()
 $app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase();$app.OpenCurrentDatabase($OutputPath,$true)
 . (Join-Path $PSScriptRoot 'verify-built-application.ps1')
 $app.CloseCurrentDatabase()
}finally{
 if($app){$app.Quit(2)}
 foreach($value in @($form,$db,$app)){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}}}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
AssertSwiftBuildInputs $swiftBuildPlan
if((Get-FileHash -LiteralPath $TemplatePath).Hash -ne $sourceHash){throw 'Template changed'}
$hash=(Get-FileHash -LiteralPath $OutputPath).Hash
@{Output=$OutputPath;SHA256=$hash;AppVersion=$swiftBuildPlan.AppVersion;SchemaVersion=$swiftBuildPlan.SchemaVersion;BuildKind='Incremental schema-27 directory merge update on verified schema-26 blank UI';Template=$TemplatePath;TemplateSHA256=$sourceHash;UpdatedModules=$updated;Inputs=$swiftBuildPlan.InputHashes;Verification=$swiftBuildVerification;Status='Source/form verified; workflow tests required'}|ConvertTo-Json -Depth 9|Set-Content ($OutputPath+'.build.json')
Write-Output ('PASS: updated '+($updated -join ', ')+'; verified '+$swiftBuildVerification.SourceModules+' installed modules and '+$swiftBuildVerification.Forms+' forms')
