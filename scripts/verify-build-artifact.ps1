param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$RequireBlank,[switch]$ForceNative,[ValidateSet('core','extended')][string]$Tier='core')
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
$swiftBuildPlan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
$artifactSnapshot=GetSwiftWorkflowArtifact $BuildPath
$reuseBuildVerification=(-not $ForceNative -and (TestSwiftBuildReceipt $artifactSnapshot $swiftBuildPlan))
$dao=$null;$db=$null;$row=$null;$app=$null
try {
    # Inspect the original through read-only DAO; execute code only in our copy.
    $dao=New-Object -ComObject DAO.DBEngine.120
    $db=$dao.OpenDatabase($artifactSnapshot.Path,$false,$true)
    $row=$db.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4)
    if($row.EOF -or $row.Fields.Item('AppVersion').Value -ne $swiftBuildPlan.AppVersion -or $row.Fields.Item('SchemaVersion').Value -ne $swiftBuildPlan.SchemaVersion){throw 'Artifact identity differs from current source'}
    $row.Close();$row=$null
    if($RequireBlank){
        # TraitDefinition is reviewed application reference data, never user data.
        $catalogue=Get-Content (Join-Path $root 'src/trait-catalogue.json') -Raw|ConvertFrom-Json
        $row=$db.OpenRecordset('SELECT Count(*) FROM TraitDefinition',4)
        if($row.Fields.Item(0).Value -ne 147){throw 'Blank trait reference inventory differs'}
        $row.Close();$row=$null
        foreach($trait in $catalogue.Traits){foreach($mark in @('','NOB','1.0','2.0','3.0','4.0','5.0')){
            $storedMark=if($mark -eq ''){'GENERAL'}else{$mark}
            $row=$db.OpenRecordset("SELECT * FROM TraitDefinition WHERE FormKind='"+$trait.FormKind+"' AND BlockNumber="+$trait.Block+" AND Mark='"+$storedMark+"' AND CatalogueVersion=1",4)
            $description=if($mark -eq ''){$trait.GeneralDescription}else{$trait.Marks.$mark}
            if($row.EOF){throw 'Blank is missing a reviewed trait definition'}
            foreach($pair in @(@('TraitTitle',$trait.Title),@('DescriptionText',$description),@('SourceName',$catalogue.Source),@('SourceSHA256',$catalogue.SourceSHA256),@('SourceTraitID',$trait.SourceTraitID))){
                $actual=$row.Fields.Item($pair[0]).Value
                if($null -eq $pair[1]){if($actual -isnot [DBNull]){throw 'Blank reference Null changed'}}elseif([string]$actual -cne [string]$pair[1]){throw ('Blank reference differs: '+$pair[0])}
            }
            $row.Close();$row=$null
        }}
        foreach($table in $db.TableDefs){
            if($table.Name -like 'MSys*' -or $table.Name -in @('AppMetadata','TraitDefinition')){continue}
            $row=$db.OpenRecordset(('SELECT TOP 1 * FROM ['+$table.Name.Replace(']',']]')+']'),4)
            if(-not $row.EOF){throw ('Blank artifact contains rows in '+$table.Name)}
            $row.Close();$row=$null
        }
    }
    $db.Close();$db=$null
    if($reuseBuildVerification){
        Write-Output ('Reusing exact-artifact compile/source/binding verification: '+$artifactSnapshot.Path+'.build.json')
    } else {
        $folder=Join-Path $root ('build/tests/artifact-'+[guid]::NewGuid())
        [void](New-Item -ItemType Directory -Path $folder)
        $target=Join-Path $folder 'verification.accdb'
        Copy-Item -LiteralPath $artifactSnapshot.Path -Destination $target
        $app=New-Object -ComObject Access.Application
        $app.AutomationSecurity=1
        $app.OpenCurrentDatabase($target,$true)
        foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
        $app.DoCmd.RunCommand(126)
        . (Join-Path $PSScriptRoot 'verify-built-application.ps1')
        $app.CloseCurrentDatabase()
    }
} finally {
    if($row){$row.Close()}
    if($db){$db.Close()}
    if($app){$app.Quit(2)}
    foreach($value in @($row,$db,$dao,$f,$checkDB,$meta,$app)){
        if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}}
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((GetSwiftWorkflowArtifact $artifactSnapshot.Path).SHA256 -ne $artifactSnapshot.SHA256){throw 'Artifact verification changed its source'}
AssertSwiftBuildInputs $swiftBuildPlan
Write-Output ('PASS: current artifact identity, compile, source and bindings; blank required: '+$RequireBlank.IsPresent+'; build verification reused: '+$reuseBuildVerification)
# -Tier extended adds installed property bindings, retired-route cleanup and the
# finalizer case (scripts/cases), each in its own process on its own copy.
if($Tier -eq 'extended'){
    . (Join-Path $PSScriptRoot 'suite-common.ps1')
    foreach($case in @('test-build-property-bindings','test-build-cleanup','test-build-finalization')){
        Invoke-CaseScript $case @('-BuildPath',$artifactSnapshot.Path)
    }
    Complete-SuiteRun 'artifact-contracts' $artifactSnapshot.Path $artifactSnapshot.SHA256
}
