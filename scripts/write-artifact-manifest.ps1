param(
 [Parameter(Mandatory=$true)][string]$ArtifactPath,
 [Parameter(Mandatory=$true)][string]$ManifestPath,
 [Parameter(Mandatory=$true)][string]$BuildRevision,
 [switch]$RequireBlank
)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$ArtifactPath=(Resolve-Path -LiteralPath $ArtifactPath).Path
$ManifestPath=[IO.Path]::GetFullPath($ManifestPath)
if(Test-Path -LiteralPath $ManifestPath){throw 'Refusing to overwrite an existing manifest.'}
$before=(Get-FileHash -LiteralPath $ArtifactPath -Algorithm SHA256).Hash
$engine=New-Object -ComObject DAO.DBEngine.120
$db=$null
$counts=[ordered]@{}
try {
 $db=$engine.OpenDatabase($ArtifactPath,$false,$true)
 foreach($table in $db.TableDefs){
  if($table.Name -like 'MSys*'){continue}
  if($table.Connect){throw 'Linked tables cannot establish a self-contained artifact.'}
  if($table.Name -eq 'AppMetadata'){continue}
  $r=$db.OpenRecordset('SELECT Count(*) FROM ['+$table.Name.Replace(']',']]')+']',4)
  try{$counts[$table.Name]=[long]$r.Fields.Item(0).Value}finally{$r.Close()}
 }
 if(-not $counts.Contains('Report') -or -not $counts.Contains('Workspace')){throw 'Missing required workspace tables.'}
 $r=$db.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4)
 try{
  if($r.EOF){throw 'Missing build metadata.'}
  $version=$r.Fields.Item('AppVersion').Value;$schema=$r.Fields.Item('SchemaVersion').Value
  $r.MoveNext();if(-not $r.EOF){throw 'Ambiguous build metadata.'}
 }finally{$r.Close()}
}finally{
 if($null -ne $db){$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)}
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)
}
$after=(Get-FileHash -LiteralPath $ArtifactPath -Algorithm SHA256).Hash
if($after -ne $before){throw 'Artifact changed during inspection; no manifest written.'}
$empty=@($counts.Values|Where-Object {$_ -ne 0}).Count -eq 0
if($RequireBlank -and -not $empty){throw 'Artifact contains data; preserve it as a workspace. No blank manifest written.'}
$checkpoint=(& git -C $root rev-parse HEAD)
if($LASTEXITCODE -ne 0){throw 'Cannot identify source checkpoint.'}
$sourceFiles=@(Get-ChildItem (Join-Path $root 'src') -File | Sort-Object Name | ForEach-Object {
 [ordered]@{Path=('src/'+$_.Name);SHA256=(Get-FileHash -LiteralPath $_.FullName).Hash}
})
$manifest=[ordered]@{
 Status='Content inventory only; not release acceptance'
 BuildRevision=$BuildRevision
 SourceCheckpoint=$checkpoint
 SourceDirty=([bool](& git -C $root status --porcelain))
 SourceFiles=$sourceFiles
 InstalledSourceAlignment='Not established by this inventory'
 AppVersion=$version;SchemaVersion=$schema
 Artifact=(Split-Path $ArtifactPath -Leaf);SHA256=$after
 InspectedAtUTC=[DateTime]::UtcNow.ToString('o')
 WorkspaceState=$(if($empty){'Empty'}else{'Populated; synthetic provenance not inferred'})
 UserDataCounts=$counts
}
$json=$manifest|ConvertTo-Json -Depth 6
$stream=[IO.File]::Open($ManifestPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
try{$bytes=[Text.Encoding]::UTF8.GetBytes($json);$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}
Write-Output ('PASS: read-only artifact inventory: '+$ManifestPath)
