param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[switch]$PrepareOnly,[switch]$VerifyViews,[switch]$InstallViewContracts)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$folder=Join-Path $root ('build/tests/diagnostics-contracts-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $folder)
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $folder 'candidate.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null;$primary=$null;$cleanup=@()
try{
 Write-Output ('Controller: diagnostics/contracts runner; database: '+$scratch)
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($view in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$view,2)}
 if($InstallSource){
  $db=$app.CurrentDb()
  if(@($db.TableDefs|ForEach-Object {$_.Name}) -notcontains 'TraitDefinition'){
   $schema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw|ConvertFrom-Json
   $db.Execute(@($schema.SQL|Where-Object {$_ -like 'CREATE TABLE TraitDefinition *'})[0],128)
   & (Join-Path (Split-Path $PSScriptRoot -Parent) 'write-trait-catalogue.ps1') -SourcePath (Join-Path $root 'src/trait-catalogue.json') -OutputPath (Join-Path $folder 'modTraitCatalogue.bas') -Application $app
  }else{
   & (Join-Path (Split-Path $PSScriptRoot -Parent) 'write-trait-catalogue.ps1') -SourcePath (Join-Path $root 'src/trait-catalogue.json') -OutputPath (Join-Path $folder 'modTraitCatalogue.bas')
  }
  $app.LoadFromText(5,'modTraitCatalogue',(Join-Path $folder 'modTraitCatalogue.bas'))
  foreach($module in Get-ChildItem (Join-Path $root 'src') -Filter '*.bas'){$app.LoadFromText(5,$module.BaseName,$module.FullName)}
  $db.Execute("UPDATE AppMetadata SET AppVersion='0.33.0-dev',SchemaVersion=31",128);$db.Close();$db=$null
 }
 if($InstallViewContracts){$app.LoadFromText(5,'modUX',(Join-Path $root 'src/modUX.bas'))}
 $app.LoadFromText(5,'DiagnosticsContractsHarness',(Join-Path $root 'tests/DiagnosticsContractsHarness.bas'))
 $app.DoCmd.RunCommand(126)
 if($VerifyViews){
  . (Join-Path (Split-Path $PSScriptRoot -Parent) 'build-plan.ps1')
  $swiftBuildPlan=GetSwiftBuildPlan $root (Join-Path (Split-Path $PSScriptRoot -Parent) 'build-manifest.json')
  . (Join-Path (Split-Path $PSScriptRoot -Parent) 'verify-built-application.ps1')
  $nativeViewFaults=0
  foreach($fault in @('renamed-control','wrong-callback')){
   $app.DoCmd.OpenForm('frmEditor',1);$editor=$app.Forms.Item('frmEditor');$save=$editor.Controls.Item('cmdSave');$originalBinding=[string]$save.OnClick
   if($fault -eq 'renamed-control'){$save.Name='SyntheticRenamedSave'}else{$save.OnClick='=CloseEditor()'}
   $app.DoCmd.Close(2,'frmEditor',1)
   $rejected=$false
   try{. (Join-Path (Split-Path $PSScriptRoot -Parent) 'verify-built-application.ps1')}
   catch{
    $expected=if($fault -eq 'renamed-control'){'*Required view control missing: frmEditor.cmdSave*'}else{'*Required view binding differs: frmEditor.cmdSave.OnClick*'}
    if($_.Exception.Message -notlike $expected){throw};$rejected=$true
   }finally{
    if($app.CurrentProject.AllForms.Item('frmEditor').IsLoaded){$app.DoCmd.Close(2,'frmEditor',2)}
    $app.DoCmd.OpenForm('frmEditor',1);$editor=$app.Forms.Item('frmEditor')
    if($fault -eq 'renamed-control'){$editor.Controls.Item('SyntheticRenamedSave').Name='cmdSave'}else{$editor.Controls.Item('cmdSave').OnClick=$originalBinding}
    $app.DoCmd.Close(2,'frmEditor',1)
   }
   if(-not $rejected){throw ('Installed verifier accepted '+$fault)}
   $nativeViewFaults++;Write-Output ('PASS: installed verifier rejects '+$fault)
  }
 }
 if($PrepareOnly){$result='PASS: prepared disposable source candidate; no behavioral assertions'}else{$result=[string]$app.Run('TestDiagnosticsContracts')}
 Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 # Reproducible inventory, not a numeric-code renumbering exercise. Every site is
 # retained for review, including legitimate optional probes and local cleanup.
 $inventory=@();$codes=@()
 foreach($source in Get-ChildItem (Join-Path $root 'src') -Filter '*.bas'){
  $procedure='';$line=0
  foreach($text in Get-Content $source.FullName){
   $line++;if($text -match '^(?:Public |Private )?(?:Function|Sub) (\w+)'){$procedure=$Matches[1]}
   if($text -match 'On Error Resume Next'){$inventory+=[pscustomobject]@{Module=$source.BaseName;Procedure=$procedure;Line=$line;Statement=$text.Trim()}}
   foreach($raise in [regex]::Matches($text,'Err\.Raise\s+([^,\r\n]+)')){
    $expression=$raise.Groups[1].Value.Trim();$kind=if($expression -match '^(vbObjectError\s*\+\s*\d+|-?\d+)$'){'Literal'}else{'Forwarded'}
    $codes+=[pscustomobject]@{Module=$source.BaseName;Procedure=$procedure;Code=$expression;Kind=$kind;Line=$line}
   }
  }
 }
 $inventory|Export-Csv (Join-Path $folder 'suppression-inventory.csv') -NoTypeInformation
 $codes|Export-Csv (Join-Path $folder 'error-identities.csv') -NoTypeInformation
 $receipt=[ordered]@{Status='PASS';Assertions=$result;SourceSHA256=$hash;TestDatabase=$scratch;InstalledSource=[bool]$InstallSource;InstalledViewContracts=[bool]$InstallViewContracts;PreparedOnly=[bool]$PrepareOnly;NativeViewFaults=$nativeViewFaults;CheckerSHA256=(Get-FileHash -LiteralPath $PSCommandPath).Hash;HarnessSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'tests/DiagnosticsContractsHarness.bas')).Hash;SuppressionSites=$inventory.Count;ErrorSites=$codes.Count;ReusedCodes=@($codes|Where-Object Kind -eq 'Literal'|Group-Object Code|Where-Object Count -gt 1|Select-Object Name,Count);Scope='Native primary/cleanup failures, rollback/retry and required view bindings; source inventories include literal codes and forwarded expressions; no release acceptance'}
}catch{$primary=$_}finally{
 if($app){try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){$cleanup+='Source artifact hash changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
$receipt|ConvertTo-Json -Depth 6|Set-Content (Join-Path $folder 'verification.json')
Write-Output ('Receipt: '+(Join-Path $folder 'verification.json'))
