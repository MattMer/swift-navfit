$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$candidate=Join-Path $root ('output\Swift_NAVFIT_0.6.0-candidate-'+[guid]::NewGuid()+'.accdb')
try {
    & (Join-Path $PSScriptRoot 'build.ps1') -OutputPath $candidate
    & (Join-Path $PSScriptRoot 'test-tools.ps1') -BuildPath $candidate
    $candidate | Set-Content -LiteralPath (Join-Path $root 'build\latest-feature-candidate.txt')
} catch {
    $errorText=$_.Exception.Message
    $test=Get-ChildItem -LiteralPath (Join-Path $root 'build\tests') -Filter 'tools-*.accdb'|Sort-Object LastWriteTime -Descending|Select-Object -First 1
    if($test){
        $engine=New-Object -ComObject DAO.DBEngine.120
        $db=$engine.OpenDatabase($test.FullName,$false,$true)
        $r=$db.OpenRecordset('SELECT CompatibilityStatus FROM AppMetadata')
        Write-Output ('Last assertion: '+$r.Fields.Item(0).Value);$r.Close()
        $r=$db.OpenRecordset('SELECT OperationID,OperationKind,OperationState,Outcome FROM OperationJournal ORDER BY OperationID')
        while(-not $r.EOF){Write-Output ("Operation $($r.Fields.Item(0).Value): $($r.Fields.Item(1).Value) / $($r.Fields.Item(2).Value) / $($r.Fields.Item(3).Value)");$r.MoveNext()}
        $r.Close();$db.Close()
    }
    throw $errorText
}
