param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$RequireBlank,[switch]$ForceNative)
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
        foreach($table in $db.TableDefs){
            if($table.Name -like 'MSys*' -or $table.Name -eq 'AppMetadata'){continue}
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
