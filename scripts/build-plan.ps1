# Read-only preflight: no Access instance or output is created until this passes.
function GetSwiftBuildPlan([string]$Root,[string]$ManifestPath){
 $manifest=Get-Content -LiteralPath $ManifestPath -Raw|ConvertFrom-Json
 if($manifest.manifestVersion -ne 1){throw 'Unsupported build manifest version'}
 $paths=@($manifest.identitySource,$manifest.schemaSource)+@($manifest.sourceModules)+@($manifest.inputs)+@($manifest.stages|ForEach-Object {$_.file})+@($manifest.forms|ForEach-Object {'src/forms/'+$_+'.form.txt'})
 $rootPath=[IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar
 $hashes=[ordered]@{}
 foreach($relative in @($paths|Sort-Object -Unique)){
  if([IO.Path]::IsPathRooted($relative)){throw ('Build input must be relative: '+$relative)}
  $path=[IO.Path]::GetFullPath((Join-Path $rootPath $relative))
  if(-not $path.StartsWith($rootPath,[StringComparison]::OrdinalIgnoreCase)){throw ('Build input escapes repository: '+$relative)}
  if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw ('Missing build dependency: '+$relative)}
  $hashes[$relative]=(Get-FileHash -LiteralPath $path).Hash
 }
 $hashes[[IO.Path]::GetFullPath($ManifestPath)]=(Get-FileHash -LiteralPath $ManifestPath).Hash
 # Stages only generate resources (modules, the PDF report, properties); forms are source.
 $seen=@{};$phaseIndex=@{'resources'=0};$last=-1
 foreach($stage in $manifest.stages){
  if($seen.ContainsKey([string]$stage.id)){throw ('Duplicate build stage: '+$stage.id)}
  if(-not $phaseIndex.ContainsKey([string]$stage.phase) -or $phaseIndex[[string]$stage.phase] -lt $last){throw ('Invalid build phase order: '+$stage.id)}
  $last=$phaseIndex[[string]$stage.phase]
  foreach($dependency in $stage.requires){if(-not $seen.ContainsKey([string]$dependency)){throw ('Build stage '+$stage.id+' requires an earlier stage: '+$dependency)}}
  $seen[[string]$stage.id]=$true
 }
 $sourceNames=@($manifest.sourceModules|ForEach-Object {[IO.Path]::GetFileName($_)}|Sort-Object)
 $actualNames=@(Get-ChildItem (Join-Path $Root 'src') -Filter '*.bas'|ForEach-Object {$_.Name}|Sort-Object)
 if(Compare-Object $sourceNames $actualNames){throw 'Build manifest source-module inventory differs from src/*.bas'}
 if(@($sourceNames|Sort-Object -Unique).Count -ne $sourceNames.Count){throw 'Duplicate source module in build manifest'}
 $formNames=@($manifest.forms|ForEach-Object {[string]$_}|Sort-Object)
 $formFiles=@(Get-ChildItem (Join-Path $Root 'src/forms') -Filter '*.form.txt'|ForEach-Object {$_.Name.Substring(0,$_.Name.Length-9)}|Sort-Object)
 if(-not $formNames.Count -or (Compare-Object $formNames $formFiles)){throw 'Build manifest form inventory differs from src/forms'}
 if(@($formNames|Sort-Object -Unique).Count -ne $formNames.Count){throw 'Duplicate form in build manifest'}
 $constants=Get-Content (Join-Path $Root $manifest.identitySource) -Raw
 $appMatch=[regex]::Match($constants,'Public Const APP_VERSION As String = "([0-9A-Za-z.\-]+)"')
 $schemaMatch=[regex]::Match($constants,'Public Const SCHEMA_VERSION As Long = (\d+)')
 if(-not $appMatch.Success -or -not $schemaMatch.Success){throw 'Cannot read canonical application identity'}
 $schema=Get-Content (Join-Path $Root $manifest.schemaSource) -Raw|ConvertFrom-Json
 if([int]$schema.SchemaVersion -ne [int]$schemaMatch.Groups[1].Value){throw 'Schema definition and application constant disagree'}
 [pscustomobject]@{Root=$Root;Manifest=$manifest;AppVersion=$appMatch.Groups[1].Value;SchemaVersion=[int]$schemaMatch.Groups[1].Value;InputHashes=$hashes}
}
function AssertSwiftBuildInputs($Plan){
 foreach($relative in $Plan.InputHashes.Keys){
  $path=if([IO.Path]::IsPathRooted($relative)){$relative}else{Join-Path $Plan.Root $relative}
  if((Get-FileHash -LiteralPath $path).Hash -ne $Plan.InputHashes[$relative]){throw ('Build input changed during generation: '+$relative)}
 }
}
