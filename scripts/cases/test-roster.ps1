param([Parameter(Mandatory=$true)][string]$BuildPath)
#F02 roster snapshots and departures: PRD window, absence from a complete roster,
#partial rosters, returns, matching, rank updates and confirm/dismiss; MNA workbooks.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/roster-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
#Writes a minimal MNA-style workbook: one sheet, shared strings, rows of cells A..W.
function New-MnaWorkbook([string]$Path,[object[]]$Rows){
    Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
    $strings=New-Object System.Collections.Generic.List[string]; $index=@{}
    $sheet=New-Object System.Text.StringBuilder
    [void]$sheet.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    for($r=0;$r -lt $Rows.Count;$r++){
        [void]$sheet.Append('<row r="'+($r+1)+'">')
        $cells=@($Rows[$r])
        for($c=0;$c -lt $cells.Count;$c++){
            $v=[string]$cells[$c]; if($v -eq ''){continue}
            if(-not $index.ContainsKey($v)){$index[$v]=$strings.Count; $strings.Add($v)}
            [void]$sheet.Append('<c r="'+[char](65+$c)+($r+1)+'" t="s"><v>'+$index[$v]+'</v></c>')
        }
        [void]$sheet.Append('</row>')
    }
    [void]$sheet.Append('</sheetData></worksheet>')
    $shared='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+$strings.Count+'" uniqueCount="'+$strings.Count+'">'+(($strings|ForEach-Object{'<si><t xml:space="preserve">'+[Security.SecurityElement]::Escape($_)+'</t></si>'}) -join '')+'</sst>'
    $parts=[ordered]@{
        '[Content_Types].xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
        '_rels/.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="CUI - New Sheet" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml'=$sheet.ToString()
        'xl/sharedStrings.xml'=$shared
        'xl/styles.xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    }
    if(Test-Path -LiteralPath $Path){Remove-Item -LiteralPath $Path}
    $zip=[IO.Compression.ZipFile]::Open($Path,'Create')
    try{
        foreach($name in $parts.Keys){
            $w=New-Object IO.StreamWriter($zip.CreateEntry($name).Open(),(New-Object Text.UTF8Encoding($false)))
            $w.Write($parts[$name]); $w.Close()
        }
    } finally {$zip.Dispose()}
}
#Three synthetic MNA rosters for UIC 99999; no real names or contacts.
$mnaDir=[IO.Path]::GetFullPath((Join-Path $root ('build/tests/mna-'+[guid]::NewGuid())))
New-Item -ItemType Directory -Path $mnaDir | Out-Null
$soon=(Get-Date).AddDays(40).ToString('yyMM')
$enlisted=@('BIN','ABSC','Billet Title','UMUIC','AUIC','RBSC','ARATE','IRATE','RFAS','NEC','SNEC','ADTE','DoDID','Name','PRD','IMS','ADM','MED','OTH','Assign Type','PNEC','SNEC','TRUIC')
$officer=@('BIN','ABSC','Billet Title','UMUIC','AUIC','RBSC','ARATE','IRATE','RFAS','NOBC','SNOBC','ADTE','DoDID','Name','PRD','IMS','ADM','MED','OTH','Assign Type','PNOBC','SNOBC','TRUIC')
function Billet($bin,$umuic,$irate,$dodid,$name,$prd,$assign,$truic){@($bin,'11111','SYN BILLET',$umuic,'11111','7001',$irate,$irate,'5AAE','','','0125',$dodid,$name,$prd,'','','','',$assign,'','',$truic)}
function MnaRows($date,[bool]$withBravo,[string]$alpha){
    $rows=@(@('Search Criteria','','','','','CUI'),@("Date: Wed $date 17:35:33 EDT 2025,  UIC: 99999"),@('Enlisted Local / CAI / Vacant: 3','Enlisted CAO: 1','Enlisted IAP: 1'),@('Enlisted Local/CAI / Vacant'),$enlisted)
    $rows+=,(Billet '0000001' '99999' 'IT2' '9000000001' $alpha '4012' 'LOCAL' '99999')
    if($withBravo){$rows+=,(Billet '0000002' '99999' 'IT1' '9000000002' 'SYNBRAVO MARK ALLEN JR' '4012' 'CROSS' '88888')}
    $rows+=,@('0000003','11111','SYN BILLET','99999','11111','7001','ITC','','5AAE','','','','','','','','','','','VACANT','','','')
    $rows+=,@('Enlisted CAO'); $rows+=,$enlisted
    $rows+=,(Billet '0000004' '77777' 'ITC' '9000000003' 'SYNCHARLIE PAUL' $soon 'CROSS' '99999')
    $rows+=,@('Enlisted IAP'); $rows+=,$enlisted
    $rows+=,(Billet '' '' 'IT2' '9000000004' 'SYNDELTA ANNE MARIE' '4012' 'IAP' '99999')
    $rows+=,@('Officer Local / CAI / Vacant: 1','Officer CAO: 0','Officer IAP: 0'); $rows+=,@('Officer Local / CAI / Vacant'); $rows+=,$officer
    $rows+=,(Billet '0000005' '99999' '1815H' '9000000005' 'SYNECHO RUTH ANN' '4012' 'LOCAL' '99999')
    $rows+=,@('','','','','','CUI'); $rows+=,@('','','','','','POC: ITC SYNTHETIC PULLER, (000) 000-0000')
    $rows
}
New-MnaWorkbook (Join-Path $mnaDir 'mna-1.xlsx') (MnaRows 'Apr 02' $true 'SYNALPHA JOHN QUINCY')
New-MnaWorkbook (Join-Path $mnaDir 'mna-2.xlsx') (MnaRows 'May 05' $false 'SYNALPHA-NEW JOHN QUINCY')
New-MnaWorkbook (Join-Path $mnaDir 'mna-3.xlsx') (MnaRows 'Mar 01' $true 'SYNALPHA JOHN QUINCY')
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
    $app.LoadFromText(5,'RosterHarness',(Join-Path $root 'tests/RosterHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestRoster()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
    $result=$app.Eval("TestMNARoster('"+$mnaDir.Replace("'","''")+"')")
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}; Remove-Item -LiteralPath $mnaDir -Recurse -Force -ErrorAction SilentlyContinue}
