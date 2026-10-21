param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $root ('build/tests/catalogue-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$catalogue=Get-Content (Join-Path $root 'src/trait-catalogue.json') -Raw|ConvertFrom-Json
$app=$null
try{
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($app.Eval('TraitCatalogueHash()') -cne $catalogue.SourceSHA256){throw 'Embedded catalogue identity changed'}
 $checks=1
 foreach($trait in $catalogue.Traits){
  $prefix='("'+$trait.FormKind+'",'+$trait.Block
  if($app.Eval('CatalogueTraitTitle'+$prefix+')') -cne $trait.Title){throw 'Trait title changed'};$checks++
  if($app.Eval('CatalogueTraitDescription'+$prefix+',"")') -cne [string]$trait.GeneralDescription){throw 'Trait general description changed'};$checks++
  foreach($mark in @('NOB','1.0','2.0','3.0','4.0','5.0')){
   if($app.Eval('CatalogueTraitDescription'+$prefix+',"'+$mark+'")') -cne [string]$trait.Marks.$mark){throw 'Trait mark description changed'};$checks++
  }
 }
 Write-Output ('PASS: '+$checks+' exact embedded trait-catalogue fidelity checks')
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){throw 'Catalogue test changed source artifact'}
