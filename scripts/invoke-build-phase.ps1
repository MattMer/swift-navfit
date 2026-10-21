param([Parameter(Mandatory=$true)][string]$BuildPhase)
# Dot-source this runner in the build's generation scope: layout helpers and
# catalogue variables intentionally survive into dependent stages.
foreach($swiftBuildStage in @($swiftBuildPlan.Manifest.stages|Where-Object {$_.phase -eq $BuildPhase})){
 $swiftStageTimer=[Diagnostics.Stopwatch]::StartNew()
 Write-Output ('Build stage: '+$swiftBuildStage.id)
 try{
  . (Join-Path $root $swiftBuildStage.file)
  $swiftStageTimer.Stop()
  $swiftBuildTrace.Add([pscustomobject]@{Id=$swiftBuildStage.id;Phase=$BuildPhase;File=$swiftBuildStage.file;Seconds=$swiftStageTimer.Elapsed.TotalSeconds;Status='PASS'})
 }catch{
  throw ('Build stage '+$swiftBuildStage.id+' failed: '+$_.Exception.Message+"`n"+$_.ScriptStackTrace)
 }
}
