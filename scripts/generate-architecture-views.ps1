param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$Check)
# Generate the DoDAF-aligned architecture views that come from tracked source
# and a fresh build (docs/architecture/views): AV-2, DIV-2, DIV-3, OV-6a, SV-8,
# plus the generated blocks in AV-1, SV-7 and TECHNICAL_README.md. -Check writes
# nothing and fails when a view is stale; it also checks the OV-6b state lists
# against the schema and the TECHNICAL_README data inventory against the tables.
# Reads a disposable copy of the build; installs the standard rule packs there.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$views=Join-Path $root 'docs/architecture/views';$nl="`n"
function Cell([string]$s){return (($s -replace '\r?\n',' ') -replace '\|','\|').Trim()}
function Table($head,$rows){
 $out=@(('| '+($head -join ' | ')+' |'),('| '+(($head|ForEach-Object {'---'}) -join ' | ')+' |'))
 foreach($r in $rows){$out+='| '+(($r|ForEach-Object {Cell ([string]$_)}) -join ' | ')+' |'}
 return $out -join $nl
}

#--- Read the build -------------------------------------------------------------
$copy=Join-Path $root ('build/tests/architecture-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $copy
$app=New-Object -ComObject Access.Application
$tables=@();$relations=@();$packs=@();$rules=@();$meta=$null;$tableRules=@{};$refs=@();$navfitHash=''
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($copy,$true)
 foreach($open in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$open,2)}
 [void]$app.Eval('InstallNavyDraftChecks()');[void]$app.Eval('InstallLegacyNavyDraftChecks()')
 $db=$app.CurrentDb()
 $r=$db.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4);$meta=@{App=[string]$r.Fields.Item(0).Value;Schema=[int]$r.Fields.Item(1).Value};$r.Close()
 $types=@{1='Yes/No';3='Integer';4='Long';6='Single';7='Double';8='Date/Time';10='Text';12='Long text'}
 foreach($t in $db.TableDefs){
  if($t.Name -like 'MSys*' -or $t.Name -like '~*' -or $t.Connect){continue}
  $pk=@();$unique=@()
  foreach($ix in $t.Indexes){
   $fields=@($ix.Fields|ForEach-Object {$_.Name}) -join ', '
   if($ix.Primary){$pk+=$fields}elseif($ix.Unique -and -not $ix.Foreign){$unique+=$fields}
  }
  $cols=@(foreach($f in $t.Fields){
   $kind=if($f.Attributes -band 16){'AutoNumber'}elseif($types.ContainsKey([int]$f.Type)){$types[[int]$f.Type]}else{'Type '+$f.Type}
   if([int]$f.Type -eq 10){$kind+=' ('+$f.Size+')'}
   [pscustomobject]@{Name=$f.Name;Type=$kind;Required=[bool]$f.Required}
  })
  $tables+=[pscustomobject]@{Name=$t.Name;Columns=$cols;Key=($pk -join '; ');Unique=(($unique|Sort-Object) -join '; ');Rule=[string]$t.ValidationRule}
  $tableRules[$t.Name]=[string]$t.ValidationRule
 }
 foreach($rel in $db.Relations){
  if($rel.Table -like 'MSys*'){continue}
  $pairs=@($rel.Fields|ForEach-Object {$_.ForeignName+' -> '+$rel.Table+'.'+$_.Name}) -join ', '
  $relations+=[pscustomobject]@{Name=$rel.Name;Parent=$rel.Table;Child=$rel.ForeignTable;Fields=$pairs}
 }
 $r=$db.OpenRecordset('SELECT PackID,VersionName,SourceTitle,EffectiveFrom,EffectiveTo FROM RulePack ORDER BY EffectiveFrom,PackID',4)
 while(-not $r.EOF){$packs+=[pscustomobject]@{ID=[int]$r.Fields.Item(0).Value;Name=[string]$r.Fields.Item(1).Value;Source=[string]$r.Fields.Item(2).Value;From=$r.Fields.Item(3).Value;To=$r.Fields.Item(4).Value};$r.MoveNext()};$r.Close()
 $r=$db.OpenRecordset('SELECT PackID,RuleCode,FieldKey,CheckKind,SourceLocator,Description,FormKind FROM RuleDefinition ORDER BY PackID,RuleCode',4)
 while(-not $r.EOF){$v=@(0..6|ForEach-Object {$r.Fields.Item($_).Value});$rules+=[pscustomobject]@{Pack=[int]$v[0];Code=[string]$v[1];Field=[string]$v[2];Kind=[string]$v[3];Locator=[string]$v[4];Text=[string]$v[5];Form=[string]$v[6]};$r.MoveNext()};$r.Close()
 foreach($ref in $app.References){$refs+=[pscustomobject]@{Name=[string]$ref.Name;Guid=[string]$ref.Guid;Version=([string]$ref.Major+'.'+[string]$ref.Minor);BuiltIn=[bool]$ref.BuiltIn}}
 $navfitHash=[string]$app.Eval('NAVFITTemplateHash()')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);Start-Sleep -Milliseconds 500;Remove-Item -LiteralPath $copy -ErrorAction SilentlyContinue}
$tables=@($tables|Sort-Object Name);$relations=@($relations|Sort-Object Child,Parent,Name)

#--- Tracked sources ------------------------------------------------------------
function ReadText($relative){return [IO.File]::ReadAllText((Join-Path $root $relative))}
$catalogue=ReadText 'src/field-definitions.json'|ConvertFrom-Json
$runners=(ReadText 'tests/runners.json'|ConvertFrom-Json).runners
$manifest=ReadText 'scripts/build-manifest.json'|ConvertFrom-Json
$modules=@($manifest.sourceModules).Count+@($manifest.generatedModules).Count
$forms=@($manifest.forms).Count
$references=@([regex]::Matches((ReadText 'src/modValidationReferences.bas'),'Case "(VR-[A-Z-]+)": pdfPage = (\d+): page = "([^"]+)": locator = "([^"]+)"')|ForEach-Object {[pscustomobject]@{Code=$_.Groups[1].Value;Page=$_.Groups[3].Value;Locator=$_.Groups[4].Value}})
$history=@(Get-ChildItem (Join-Path $root 'tests/historical-schemas') -Filter 'schema-*.sql'|Sort-Object Name|ForEach-Object {
 $text=Get-Content $_.FullName;$head=@{};foreach($l in $text){if($l -match '^-- (\w+): (.*)$' -and -not $head.ContainsKey($Matches[1])){$head[$Matches[1]]=$Matches[2]}}
 $names=@($text|ForEach-Object {if($_ -match '^CREATE TABLE \[?(\w+)\]?'){$Matches[1]}})
 [pscustomobject]@{Schema=$head['schema']+$(if($head['variant']){$head['variant']}else{''});Kind=$head['kind'];App=$(if($head['appVersion']){$head['appVersion']}elseif($head['release']){$head['release']}else{''});Tables=$names}
})

#SBOM inputs: Windows API declarations and late-bound COM classes in the VBA source.
$source=@(Get-ChildItem (Join-Path $root 'src') -Filter '*.bas'|Sort-Object Name|ForEach-Object {[IO.File]::ReadAllText($_.FullName)}) -join $nl
$apis=@{}
foreach($m in [regex]::Matches($source,'Declare PtrSafe Function (\w+) Lib "(\w+)"(?: Alias "(\w+)")?')){
 $lib=$m.Groups[2].Value.ToLower()+'.dll';$fn=if($m.Groups[3].Success){$m.Groups[3].Value}else{$m.Groups[1].Value}
 if(-not $apis.ContainsKey($lib)){$apis[$lib]=New-Object 'System.Collections.Generic.SortedSet[string]'};[void]$apis[$lib].Add($fn)
}
$com=@{};foreach($m in [regex]::Matches($source,'CreateObject\("([^"]+)"\)')){$com[$m.Groups[1].Value]=1+[int]$com[$m.Groups[1].Value]}
$purpose=@{
 'ADODB.Stream'='UTF-8 and binary file reading and writing (backups, PDF output, RUAD CSV, NAVFIT template)'
 'MSXML2.DOMDocument.6.0'='XML parsing and writing (PDF form data and exchange)'
 'Scripting.Dictionary'='In-memory keyed collections'
 'Scripting.FileSystemObject'='File and folder checks'
 'Shell.Application'='Opens a generated PDF in the default viewer'
}
foreach($k in $com.Keys){if(-not $purpose.ContainsKey($k)){throw ('Undocumented COM dependency '+$k+': add its purpose in generate-architecture-views.ps1')}}
$images=@(Get-ChildItem (Join-Path $root 'src/pdf-templates') -Filter '*.png'|Sort-Object Name|ForEach-Object {[pscustomobject]@{Name=$_.Name;Hash=(Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower()}})
$devTools=@(@('Windows PowerShell 5.1','Build, test and release scripts'),@('Python 3 with pikepdf','Synthetic PDF test fixtures (scripts/test-pdf-fixtures.ps1)'),@('Python 3 with reportlab and pypdf','User guide PDF rendering (scripts/build_user_guide.py)'),@('git and GitHub CLI','Source control and pre-release publishing'))

#--- Views ----------------------------------------------------------------------
$generated='<!-- Generated by scripts/generate-architecture-views.ps1 from tracked source and a fresh build. Do not edit by hand. -->'
$out=[ordered]@{}

$rows=@($catalogue.fields|Sort-Object ordinal,key|ForEach-Object {
 $opts=@($_.options|Where-Object {$_});$o=if($opts.Count -gt 6){($opts[0..5] -join ', ')+', ...'}else{$opts -join ', '}
 ,@($_.key,$_.label,$(if($_.block){$_.block}else{'-'}),$_.storageType,$o)
})
$out['AV-2.md']=@($generated,'','# AV-2 Integrated Dictionary','',
 ('Report fields from the canonical field catalogue (`src/field-definitions.json`, catalogue version '+$catalogue.catalogueVersion+'). Each field has the same key in the editor, saved revisions, recovery, NAVFIT98A mapping and PDF mapping. Terms such as ITA, PSG RSCA and stale are defined in the [user guide glossary](../../Swift_NAVFIT_User_Guide.md#192-glossary); tables are described in [DIV-3](DIV-3.md).'),'',
 (Table @('Key','Label','Block','Storage','Choices') $rows)) -join $nl

$out['DIV-2.md']=@($generated,'','# DIV-2 Logical Data Model','',
 ('Entities and enforced relationships in schema '+$meta.Schema+'. '+$tables.Count+' entities; '+$relations.Count+' relationships. Some links are deliberately not enforced as relationships (Access allows 32 indexes per table, and saves replace `ReportHead` rows); those are checked in code and listed as fields in [DIV-3](DIV-3.md).'),'',
 '```mermaid','erDiagram',
 (@($relations|ForEach-Object {'  '+$_.Parent+' ||--o{ '+$_.Child+' : "'+$_.Name+'"'}) -join $nl),'```','',
 (Table @('Child','Parent','Fields','Relationship') @($relations|ForEach-Object {,@($_.Child,$_.Parent,$_.Fields,$_.Name)}))) -join $nl

$div3=@($generated,'','# DIV-3 Physical Data Model','',('Microsoft Access (ACE) tables in schema '+$meta.Schema+', application '+$meta.App+'. Types are Access field types; "required" means the field may not be empty.'),'')
foreach($t in $tables){
 $div3+='## '+$t.Name;$div3+=''
 if($t.Key){$div3+='- **Primary key:** '+$t.Key}
 if($t.Unique){$div3+='- **Unique:** '+$t.Unique}
 if($t.Rule){$div3+='- **Validation rule:** `'+(($t.Rule -replace '`',"'") -replace '\r?\n',' ')+'`'}
 $div3+='';$div3+=(Table @('Column','Type','Required') @($t.Columns|ForEach-Object {,@($_.Name,$_.Type,$(if($_.Required){'yes'}else{''}))}));$div3+=''
}
$out['DIV-3.md']=($div3 -join $nl).TrimEnd()

$ov6a=@($generated,'','# OV-6a Operational Rules','','Rules come from two sources. Each finding a user sees names its rule and citation.','','## Effective-dated rule packs','',
 'Installed into every workspace; a report is checked against the pack in effect for its period.','')
foreach($p in $packs){
 $to=if($p.To -is [DateTime]){$p.To.ToString('yyyy-MM-dd')}else{'open'};$from=if($p.From -is [DateTime]){$p.From.ToString('yyyy-MM-dd')}else{'-'}
 $ov6a+='### '+$p.Name;$ov6a+='';$ov6a+=('- **Source:** '+$p.Source);$ov6a+=('- **Effective:** '+$from+' to '+$to);$ov6a+=''
 $ov6a+=(Table @('Rule','Field','Check','Applies to','Citation','Rule text') @($rules|Where-Object {$_.Pack -eq $p.ID}|ForEach-Object {,@($_.Code,$_.Field,$_.Kind,$(if($_.Form){$_.Form}else{'all'}),$_.Locator,$_.Text)}));$ov6a+=''
}
$ov6a+='## Expanded rules with verified citations';$ov6a+='';$ov6a+='From `src/modValidationReferences.bas`: BUPERSINST 1610.10H CH-2 (26 May 2026 compilation), each matched against the fingerprinted source PDF.';$ov6a+=''
$ov6a+=(Table @('Rule','Printed page','Locator') @($references|Sort-Object Code|ForEach-Object {,@($_.Code,$_.Page,$_.Locator)}))
$out['OV-6a.md']=($ov6a -join $nl).TrimEnd()

$sv8=@($generated,'','# SV-8 Systems Evolution','',('Schema history from `tests/historical-schemas/` and the current build. Each release imports workspaces from schema 35 (release 0.49.0) onward through the production import (`test-workspace-migration`); release 0.55.0 also imports schemas 1-34 and bridges older workspaces. Changes are additive: tables are added, never removed or renamed.'),'')
$prev=@();$rows=@()
foreach($h in $history){$added=@($h.Tables|Where-Object {$prev -notcontains $_}|Sort-Object);$rows+=,@($h.Schema,$h.Kind,$h.App,$h.Tables.Count,$(if($prev.Count -eq 0){'(first captured)'}elseif($added.Count){$added -join ', '}else{'-'}));$prev=$h.Tables}
$current=@($tables.Name);$added=@($current|Where-Object {$prev -notcontains $_}|Sort-Object)
$rows+=,@([string]$meta.Schema,'current build',$meta.App,$current.Count,$(if($added.Count){$added -join ', '}else{'-'}))
$sv8+=(Table @('Schema','Definition','Application','Tables','Tables added') $rows)
$out['SV-8.md']=($sv8 -join $nl).TrimEnd()

$tiers=$runners|Group-Object tier|Sort-Object Name
$facts=@('| Fact | Value |','| --- | --- |',('| Application version | '+$meta.App+' |'),('| Schema version | '+$meta.Schema+' |'),('| Tables / enforced relationships | '+$tables.Count+' / '+$relations.Count+' |'),('| Report fields in the catalogue | '+@($catalogue.fields).Count+' |'),('| VBA modules / forms | '+$modules+' / '+$forms+' |'),('| Validation rule packs / pack rules / expanded rules | '+$packs.Count+' / '+$rules.Count+' / '+$references.Count+' |')) -join $nl
$counts=@('Test runners by tier (`tests/runners.json`):','',(Table @('Tier','Runners') @($tiers|ForEach-Object {,@($_.Name,$_.Count)}))) -join $nl
$sb=@($generated,'','# Software bill of materials (SBOM)','',
 ('Components Swift NAVFIT '+$meta.App+' uses at run time, from a fresh build and the VBA source. The same inventory is in CycloneDX 1.5 form in [sbom.cdx.json](sbom.cdx.json). Swift ships no third-party code libraries: everything it calls at run time is part of Windows or Microsoft Office. Swift NAVFIT itself is licensed under the Apache License 2.0 ([LICENSE](../../LICENSE)).'),'',
 '## Platform','','| Component | Supplier | Requirement |','| --- | --- | --- |','| Microsoft Access (desktop), VBA 7, ACE database engine | Microsoft | Desktop Access 2016 or later, including Microsoft 365. Access Runtime is not supported. |','',
 '## VBA project references','',(Table @('Reference','Version','GUID','Built in') @($refs|Sort-Object Name|ForEach-Object {,@($_.Name,$_.Version,$_.Guid,$(if($_.BuiltIn){'yes'}else{''}))})),'',
 '## Windows API functions (Declare)','',(Table @('Library','Functions') @($apis.Keys|Sort-Object|ForEach-Object {,@($_,(@($apis[$_]) -join ', '))})),'',
 '## Late-bound COM classes (CreateObject)','',(Table @('Class','Uses','Purpose') @($com.Keys|Sort-Object|ForEach-Object {,@($_,$com[$_],$purpose[$_])})),'',
 '## Embedded data','',(Table @('Item','Origin','SHA-256') (@(,@('NAVFIT98A blank database template','Supplied NAVFIT98A blank (FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb), checked against this hash before every export',$navfitHash.ToLower()))+@($images|ForEach-Object {,@(('src/pdf-templates/'+$_.Name),'Page image of the official NAVPERS form, used for PDF output',$_.Hash)}))),'',
 '## Development and test tools (not shipped)','',(Table @('Tool','Used for') $devTools)) -join $nl
$out['../SBOM.md']=$sb.TrimEnd()
$cdx=[ordered]@{bomFormat='CycloneDX';specVersion='1.5';version=1
 metadata=[ordered]@{component=[ordered]@{type='application';name='Swift NAVFIT';version=$meta.App;licenses=@([ordered]@{license=[ordered]@{id='Apache-2.0'}})}}
 components=@(
  @([ordered]@{type='platform';'bom-ref'='platform:access';name='Microsoft Access (desktop)';supplier=[ordered]@{name='Microsoft'};version='16.0';description='Desktop Access 2016 or later with VBA 7 and the ACE database engine; Access Runtime not supported'})+
  @($refs|Sort-Object Name|ForEach-Object {[ordered]@{type='library';'bom-ref'=('vba-ref:'+$_.Name);name=$_.Name;version=$_.Version;properties=@([ordered]@{name='vba:guid';value=$_.Guid},[ordered]@{name='vba:builtIn';value=([string]$_.BuiltIn).ToLower()})}})+
  @($apis.Keys|Sort-Object|ForEach-Object {[ordered]@{type='library';'bom-ref'=('winapi:'+$_);name=$_;supplier=[ordered]@{name='Microsoft'};description=('Windows API: '+(@($apis[$_]) -join ', '))}})+
  @($com.Keys|Sort-Object|ForEach-Object {[ordered]@{type='library';'bom-ref'=('com:'+$_);name=$_;supplier=[ordered]@{name='Microsoft'};description=$purpose[$_];properties=@([ordered]@{name='swift:uses';value=[string]$com[$_]})}})+
  @([ordered]@{type='file';'bom-ref'='data:navfit-blank';name='NAVFIT98A blank database template';hashes=@([ordered]@{alg='SHA-256';content=$navfitHash.ToLower()})})+
  @($images|ForEach-Object {[ordered]@{type='file';'bom-ref'=('data:'+$_.Name);name=('src/pdf-templates/'+$_.Name);hashes=@([ordered]@{alg='SHA-256';content=$_.Hash})}})+
  @($devTools|ForEach-Object {[ordered]@{type='application';'bom-ref'=('tool:'+$_[0]);name=$_[0];scope='excluded';description=$_[1]}})
 )}
$out['../sbom.cdx.json']=($cdx|ConvertTo-Json -Depth 10)
$migration=ReadText 'src/modMigration.bas'
$sourceRange=[regex]::Match($migration,'If ver < (\d+) Or ver > (\d+) Then');$sourceMin=$sourceRange.Groups[1].Value;$sourceMax=$sourceRange.Groups[2].Value
$targetRange=[regex]::Match($migration,'targetVersion < (\d+) Or targetVersion > (\d+) Then')
if(-not $sourceMax -or -not $targetRange.Success){throw 'Importer schema range not found in src/modMigration.bas'}
$readmeFacts=@('| Fact | Value |','| --- | --- |',
 ('| Application version | '+$meta.App+' (`APP_VERSION` in `src/modSwift.bas`) |'),
 ('| Schema version | '+$meta.Schema+' |'),
 ('| Imports accepted | Workspaces of schema '+$sourceMin+' to '+$sourceMax+', into a target of schema '+$targetRange.Groups[1].Value+' to '+$targetRange.Groups[2].Value+' |'),
 ('| Tables / enforced relationships | '+$tables.Count+' / '+$relations.Count+' |'),
 ('| VBA modules / forms | '+$modules+' / '+$forms+' |'),
 ('| Report fields in the catalogue | '+@($catalogue.fields).Count+' |'),
 ('| Test suites by tier | '+(@($tiers|ForEach-Object {[string]$_.Count+' '+$_.Name}) -join ', ')+' |'),
 ('| Test cases / VBA test harnesses | '+@(Get-ChildItem (Join-Path $root 'scripts/cases') -Filter '*.ps1').Count+' / '+@(Get-ChildItem (Join-Path $root 'tests') -Filter '*.bas').Count+' |')) -join $nl
$blocks=@{'AV-1.md'=@{facts=$facts};'SV-7.md'=@{counts=$counts};'../../../TECHNICAL_README.md'=@{'readme-facts'=$readmeFacts}}

#--- Write or check -------------------------------------------------------------
$stale=@()
foreach($name in $out.Keys){
 $path=Join-Path $views $name;$want=($out[$name] -replace "`r`n","`n")+$nl
 $have=if(Test-Path $path){[IO.File]::ReadAllText($path) -replace "`r`n","`n"}else{''}
 if($have -ne $want){if($Check){$stale+=$name}else{[IO.File]::WriteAllText($path,$want)}}
}
foreach($name in $blocks.Keys){
 $path=Join-Path $views $name;$text=[IO.File]::ReadAllText($path) -replace "`r`n","`n";$new=$text
 foreach($key in $blocks[$name].Keys){
  $pattern='(?s)(<!-- generated:'+$key+' -->\n).*?(<!-- end generated:'+$key+' -->)'
  if($new -notmatch $pattern){throw ('Missing generated block '+$key+' in '+$name)}
  $new=[regex]::Replace($new,$pattern,{param($m) $m.Groups[1].Value+$blocks[$name][$key]+$nl+$m.Groups[2].Value})
 }
 if($new -ne $text){if($Check){$stale+=$name}else{[IO.File]::WriteAllText($path,$new)}}
}
#OV-6b: each declared state list must match the values the schema enforces.
$ov6b=[IO.File]::ReadAllText((Join-Path $views 'OV-6b.md'));$checked=0
foreach($m in [regex]::Matches($ov6b,'<!-- states:(\w+)\.(\w+) ([^>]+) -->')){
 $rule=$tableRules[$m.Groups[1].Value];if(-not $rule){throw ('OV-6b names a table with no validation rule: '+$m.Groups[1].Value)}
 $in=[regex]::Match($rule,'\['+$m.Groups[2].Value+'\] In \(([^)]*)\)');if(-not $in.Success){throw ('OV-6b state field not constrained: '+$m.Groups[1].Value+'.'+$m.Groups[2].Value)}
 $enforced=@([regex]::Matches($in.Groups[1].Value,"'([^']*)'")|ForEach-Object {$_.Groups[1].Value}|Sort-Object)
 $declared=@($m.Groups[3].Value -split ','|Sort-Object)
 if(($enforced -join '|') -ne ($declared -join '|')){throw ('OV-6b states for '+$m.Groups[1].Value+'.'+$m.Groups[2].Value+' are '+($declared -join ', ')+'; the schema enforces '+($enforced -join ', '))}
 $checked++
}
if($checked -lt 2){throw 'OV-6b state lists missing'}
#TECHNICAL_README section 5.2 must classify exactly the build's tables.
$readme=[IO.File]::ReadAllText((Join-Path $root 'TECHNICAL_README.md'))
$inventory=[regex]::Match($readme,'(?s)<!-- data-inventory -->(.*?)<!-- end data-inventory -->')
if(-not $inventory.Success){throw 'TECHNICAL_README.md has no data inventory block'}
$listed=@([regex]::Matches($inventory.Groups[1].Value,'`(\w+)`')|ForEach-Object {$_.Groups[1].Value}|Sort-Object -Unique)
$unlisted=@($tables.Name|Where-Object {$listed -notcontains $_});$gone=@($listed|Where-Object {@($tables.Name) -notcontains $_})
if($unlisted.Count -or $gone.Count){throw ('TECHNICAL_README.md data inventory (section 5.2) does not match the build. Classify: '+($unlisted -join ', ')+'; remove: '+($gone -join ', '))}
if($stale.Count){throw ('Architecture views are stale; run scripts/generate-architecture-views.ps1 -BuildPath <fresh build>: '+($stale -join ', '))}
Write-Output ('PASS: '+$(if($Check){'checked'}else{'generated'})+' '+($out.Count+$blocks.Count)+' architecture views; '+$checked+' OV-6b state lists match the schema; schema '+$meta.Schema+', '+$tables.Count+' tables, '+$rules.Count+' pack rules')
