param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$OutputPath)
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
function VBText([string]$value){
    '"'+$value.Replace('"','""').Replace("`r",'" & vbCr & "').Replace("`n",'" & vbLf & "').Replace("`t",'" & vbTab & "')+'"'
}
function AssignText([string]$functionName,[string]$value){
    if([string]::IsNullOrEmpty($value)){return "        $functionName = `"`""}
    $lines=@()
    for($offset=0;$offset -lt $value.Length;$offset+=120){
        $chunk=$value.Substring($offset,[Math]::Min(120,$value.Length-$offset))
        $prefix=if($offset -eq 0){''}else{"$functionName & "}
        $lines+="        $functionName = "+$prefix+(VBText $chunk)
    }
    $lines
}
$lines=@('Attribute VB_Name = "modTraitCatalogue"','Option Compare Binary','Option Explicit','''Generated from the checked-in catalogue; source values remain data, never executable expressions.',
    'Public Function TraitCatalogueHash() As String',('    TraitCatalogueHash = "'+$catalogue.SourceSHA256+'"'),'End Function',
    'Public Function CatalogueTraitTitle(ByVal kind As String, ByVal block As Long) As String','    Select Case kind & ":" & block')
foreach($trait in $catalogue.Traits){
    $lines+='    Case '+(VBText ($trait.FormKind+':'+$trait.Block))
    $lines+=AssignText 'CatalogueTraitTitle' $trait.Title
}
$lines+=@('    End Select','End Function','Public Function CatalogueTraitDescription(ByVal kind As String, ByVal block As Long, ByVal mark As String) As String','    Select Case kind & ":" & block & ":" & mark')
foreach($trait in $catalogue.Traits){
    $lines+='    Case '+(VBText ($trait.FormKind+':'+$trait.Block+':'))
    $lines+=AssignText 'CatalogueTraitDescription' $trait.GeneralDescription
    foreach($mark in @('NOB','1.0','2.0','3.0','4.0','5.0')){
        $lines+='    Case '+(VBText ($trait.FormKind+':'+$trait.Block+':'+$mark))
        $lines+=AssignText 'CatalogueTraitDescription' $trait.Marks.$mark
    }
}
$lines+=@('    End Select','End Function')
[IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputPath),($lines -join "`r`n")+"`r`n",[Text.UTF8Encoding]::new($false))
