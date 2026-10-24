param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$OutputPath,[object]$Application)
$ErrorActionPreference='Stop'
$catalogue=Get-Content -LiteralPath $SourcePath -Raw | ConvertFrom-Json
if($catalogue.SourceSHA256 -notmatch '^[0-9A-F]{64}$' -or $catalogue.Traits.Count -ne 21){throw 'Invalid trait catalogue identity or count'}
$seen=@{}
foreach($trait in $catalogue.Traits){
    $key=$trait.FormKind+':'+$trait.Block
    if($trait.FormKind -notin @('EVAL','CHIEFEVAL','FITREP') -or $trait.Block -notin (33..39) -or $seen.ContainsKey($key)){throw 'Invalid or duplicate trait definition'}
    $seen[$key]=$true
    $marks=@($trait.Marks.PSObject.Properties.Name | Sort-Object)
    if(($marks -join ',') -ne '1.0,2.0,3.0,4.0,5.0,NOB'){throw 'Invalid trait marks'}
}
# Only identity and thin accessors are generated. Official wording is runtime data.
$lines=@('Attribute VB_Name = "modTraitCatalogue"','Option Compare Binary','Option Explicit',
 'Public Const TRAIT_CATALOGUE_VERSION As Long = 1',
 'Public Function TraitCatalogueHash() As String',(' TraitCatalogueHash = "'+$catalogue.SourceSHA256+'"'),'End Function',
 'Public Function CatalogueTraitTitle(ByVal kind As String, ByVal block As Long) As String',
 ' CatalogueTraitTitle = TraitDefinitionValue(kind, block, "", True)','End Function',
 'Public Function CatalogueTraitDescription(ByVal kind As String, ByVal block As Long, ByVal mark As String) As String',
 ' CatalogueTraitDescription = TraitDefinitionValue(kind, block, mark, False)','End Function')
[IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputPath),($lines -join "`r`n")+"`r`n",[Text.UTF8Encoding]::new($false))
if($null -eq $Application){return}
$db=$null;$query=$null;$workspace=$null;$transaction=$false
try{
 $db=$Application.CurrentDb();$workspace=$Application.DBEngine.Workspaces.Item(0)
 $db.TableDefs.Item('TraitDefinition').ValidationRule="[FormKind] In ('EVAL','CHIEFEVAL','FITREP') And [BlockNumber] Between 33 And 39 And [Mark] In ('GENERAL','NOB','1.0','2.0','3.0','4.0','5.0') And [CatalogueVersion]>0"
 $workspace.BeginTrans();$transaction=$true
 $query=$db.OpenRecordset('TraitDefinition',2)
 $fieldNames=@('FormKind','BlockNumber','Mark','CatalogueVersion','TraitTitle','DescriptionText','SourceName','SourceSHA256','SourceTraitID')
 foreach($trait in $catalogue.Traits){foreach($mark in @('','NOB','1.0','2.0','3.0','4.0','5.0')){
  $description=if($mark -eq ''){$trait.GeneralDescription}else{$trait.Marks.$mark}
  $storedMark=if($mark -eq ''){'GENERAL'}else{$mark}
  $values=@($trait.FormKind,[int]$trait.Block,$storedMark,1,$trait.Title,$description,$catalogue.Source,$catalogue.SourceSHA256,[int]$trait.SourceTraitID)
  $query.AddNew()
  for($i=0;$i -lt $values.Count;$i++){
   if($null -ne $values[$i]){$query.Fields.Item([string]$fieldNames[$i]).Value=[string]$values[$i]}
  }
  $query.Update()
 }}
 $workspace.CommitTrans();$transaction=$false
}finally{
 if($transaction){$workspace.Rollback()}
 if($query){$query.Close()};if($db){$db.Close()}
 foreach($owned in @($query,$db,$workspace)){if($null -ne $owned -and [Runtime.InteropServices.Marshal]::IsComObject($owned)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned)}}
}
