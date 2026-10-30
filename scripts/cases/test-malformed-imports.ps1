param([Parameter(Mandatory=$true)][string]$BuildPath)
# Malformed-input tests for the PDF, RUAD CSV and NAVFIT98A importers (threat
# model recommendation 2). Fixtures are generated here and by the harness from
# Swift's embedded NAVFIT98A blank; nothing outside the repository is needed.
# Each malformed case must be rejected with a message within 20 seconds, with
# no rows written, the source unchanged, importer state cleared and no temp
# files left. Control cases must load.
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$run=Join-Path $root ('build/tests/malformed-'+[guid]::NewGuid());New-Item -ItemType Directory -Path $run|Out-Null
$latin=[Text.Encoding]::GetEncoding(28591)
function Bytes([string]$s){return ,$latin.GetBytes($s)}
function Join-Bytes([object[]]$parts){$m=New-Object IO.MemoryStream;foreach($p in $parts){$m.Write($p,0,$p.Length)};return ,$m.ToArray()}
function Save([string]$name,[byte[]]$data){[IO.File]::WriteAllBytes((Join-Path $run $name),$data);return $name}
function ZeroBomb([int]$megabytes){
 #Compress zeros chunk by chunk; the Adler-32 of n zeros is 1 + (n mod 65521) * 65536.
 $m=New-Object IO.MemoryStream;$d=New-Object IO.Compression.DeflateStream($m,[IO.Compression.CompressionLevel]::Optimal,$true);$chunk=New-Object byte[] 1048576
 for($i=0;$i -lt $megabytes;$i++){$d.Write($chunk,0,$chunk.Length)};$d.Dispose()
 [uint64]$n=[uint64]$megabytes*1048576;[uint32]$b=[uint32]($n%65521);[uint32]$a=1
 return ,(Join-Bytes @([byte[]](0x78,0x9C),$m.ToArray(),[byte[]](($b -shr 8),($b -band 255),($a -shr 8),($a -band 255))))
}
function Stream([byte[]]$data,[string]$extra=''){return ,(Join-Bytes @((Bytes ('<< /Length '+$data.Length+$extra+' >>'+"`nstream`n")),$data,(Bytes "`nendstream")))}
#A classic PDF with an XFA template and datasets packet; $objects are object bodies (bytes), numbered from 1.
function Pdf([object[]]$objects,[string]$trailerExtra='',[string]$startxref=''){
 $out=New-Object IO.MemoryStream;$offsets=@()
 $h=Bytes "%PDF-1.7`n";$out.Write($h,0,$h.Length)
 for($i=0;$i -lt $objects.Count;$i++){$offsets+=$out.Position;$p=Join-Bytes @((Bytes (''+($i+1)+" 0 obj`n")),[byte[]]$objects[$i],(Bytes "`nendobj`n"));$out.Write($p,0,$p.Length)}
 $xref=$out.Position;$t="xref`n0 "+($objects.Count+1)+"`n0000000000 65535 f `n"
 foreach($o in $offsets){$t+=('{0:D10} 00000 n ' -f $o)+"`n"}
 $sx=if($startxref){$startxref}else{[string]$xref}
 $t+='trailer'+"`n<< /Size "+($objects.Count+1)+' /Root 1 0 R'+$trailerExtra+" >>`nstartxref`n"+$sx+"`n%%EOF`n"
 $p=Bytes $t;$out.Write($p,0,$p.Length);return ,$out.ToArray()
}
$template=Bytes '<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/"><subform name="form1"/></template>'
$datasets=Bytes '<xfa:datasets xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"><xfa:data><form1><Name>SYNTHETIC, MEMBER</Name></form1></xfa:data></xfa:datasets>'
function Base([string]$acro='<< /XFA [(template) 4 0 R (datasets) 5 0 R] >>',[object]$templateStream=$null,[object]$dataStream=$null){
 if($null -eq $templateStream){$templateStream=Stream $template};if($null -eq $dataStream){$dataStream=Stream $datasets}
 return @((Bytes '<< /Type /Catalog /Pages 2 0 R /AcroForm 3 0 R >>'),(Bytes '<< /Type /Pages /Count 0 /Kids [] >>'),(Bytes $acro),$templateStream,$dataStream)
}
$rand=New-Object Random 1610
function Random-Bytes([int]$n,[string]$prefix=''){$b=New-Object byte[] $n;$rand.NextBytes($b);$p=Bytes $prefix;[Array]::Copy($p,$b,$p.Length);return ,$b}

$cases=[Collections.Generic.List[object]]::new()
function Add([string]$kind,[string]$file,[string]$expect,[string]$field='',[string]$literal=''){$cases.Add([pscustomobject]@{Kind=$kind;File=$file;Expect=$expect;Field=$field;Literal=$literal})}

#--- PDF ---------------------------------------------------------------------------
$good=Pdf (Base)
Add PDF (Save 'pdf-control.pdf' $good) ''
Add PDF (Save 'pdf-empty.pdf' (New-Object byte[] 0)) 'empty or exceeds'
Add PDF (Save 'pdf-not-pdf.pdf' (Random-Bytes 4096)) 'Not a supported PDF'
Add PDF (Save 'pdf-header-only.pdf' (Bytes "%PDF-1.7`n%%EOF`n`n`n`n`n`n")) 'cross-reference pointer'
Add PDF (Save 'pdf-random-body.pdf' (Random-Bytes 65536 "%PDF-1.7`n")) '*'
$text=$latin.GetString($good)
Add PDF (Save 'pdf-startxref-beyond.pdf' (Bytes ($text -replace 'startxref\n\d+',('startxref'+"`n"+'99999999')))) 'cross-reference pointer'
Add PDF (Save 'pdf-startxref-text.pdf' (Bytes ($text -replace 'startxref\n\d+',('startxref'+"`n"+'abc')))) 'cross-reference pointer'
Add PDF (Save 'pdf-truncated.pdf' ($good[0..([int]($good.Length/2))])) '*'
Add PDF (Save 'pdf-trailing-junk.pdf' (Join-Bytes @($good,(Bytes 'junk')))) 'final PDF'
$xrefAt=$text.LastIndexOf("`nxref`n")+1
Add PDF (Save 'pdf-xref-cycle.pdf' (Pdf (Base) (' /Prev '+$xrefAt))) 'revision chain'
#Objects 6 and 7 refer to each other; the reader must not follow the loop.
Add PDF (Save 'pdf-reference-loop.pdf' (Pdf (@(Base '<< /XFA 6 0 R >>')+@((Bytes '7 0 R'),(Bytes '6 0 R'))))) 'Unsupported XFA layout'
Add PDF (Save 'pdf-deep-nesting.pdf' (Pdf (Base ('<< /XFA '+('['*200)+(']'*200)+' >>')))) 'structure limit'
Add PDF (Save 'pdf-deflate-bomb.pdf' (Pdf (Base -dataStream (Stream (ZeroBomb 64) ' /Filter /FlateDecode')))) 'size limit'
Add PDF (Save 'pdf-length-huge.pdf' (Pdf (Base -dataStream (Join-Bytes @((Bytes "<< /Length 999999999 >>`nstream`n"),$datasets,(Bytes "`nendstream")))))) '*'
Add PDF (Save 'pdf-length-negative.pdf' (Pdf (Base -dataStream (Join-Bytes @((Bytes "<< /Length -5 >>`nstream`n"),$datasets,(Bytes "`nendstream")))))) '*'
Add PDF (Save 'pdf-no-endstream.pdf' (Pdf (Base -dataStream (Join-Bytes @((Bytes ('<< /Length '+($datasets.Length+40)+" >>`nstream`n")),$datasets))))) '*'
$pairs=(1..130|ForEach-Object {"(p$_) 5 0 R"}) -join ' '
Add PDF (Save 'pdf-packets-many.pdf' (Pdf (Base ('<< /XFA ['+$pairs+'] >>')))) 'XFA packet array'
Add PDF (Save 'pdf-packet-not-stream.pdf' (Pdf (Base '<< /XFA [(template) 4 0 R (datasets) 2 0 R] >>'))) 'XFA packet reference'
Add PDF (Save 'pdf-unterminated-string.pdf' (Pdf (Base '<< /XFA [(template 4 0 R (datasets) 5 0 R] >>'))) '*'
Add PDF (Save 'pdf-huge-number.pdf' (Pdf (Base '<< /XFA [(template) 99999999999999999999 0 R (datasets) 5 0 R] >>'))) '*'
Add PDF (Save 'pdf-no-xfa.pdf' (Pdf (Base '<< /Fields [] >>'))) 'No supported XFA'
$big=New-Object byte[] 30000001;$hp=Bytes '%PDF-1.7';[Array]::Copy($hp,$big,$hp.Length)
Add PDF (Save 'pdf-oversize.pdf' $big) '30 MB'

#--- RUAD CSV ----------------------------------------------------------------------
$copy=Join-Path $root ('build/tests/malformed-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $copy
$app=New-Object -ComObject Access.Application;$passed=0;$failed=@()
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($copy,$true)
 foreach($open in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$open,2)}
 $app.LoadFromText(5,'MalformedImportHarness',(Join-Path $root 'tests/MalformedImportHarness.bas'));$app.DoCmd.RunCommand(126)
 $keys=@(([string]$app.Eval('RUADKeysLine()')) -split ',')
 function Row([hashtable]$override=@{}){return (@($keys|ForEach-Object {$v=if($override.ContainsKey($_)){$override[$_]}else{'A'};'"'+($v -replace '"','""')+'"'}) -join ',')}
 $header=(@($keys|ForEach-Object {'"'+$_+'"'}) -join ',')
 $utf8=New-Object Text.UTF8Encoding($false)
 function Csv([string]$name,[string]$content){[IO.File]::WriteAllText((Join-Path $run $name),$content,$utf8);return $name}
 Add CSV (Csv 'csv-control.csv' ($header+"`r`n"+(Row @{NAME='SYNTHETIC, MEMBER'})+"`r`n")) ''
 $formula='=HYPERLINK("http://example.invalid","x")'
 Add CSV (Csv 'csv-formula.csv' ($header+"`r`n"+(Row @{NAME=$formula})+"`r`n")) '' 'NAME' $formula
 Add CSV (Csv 'csv-formula-plus.csv' ($header+"`r`n"+(Row @{NAME="+cmd|' /C calc'!A0"})+"`r`n")) '' 'NAME' "+cmd|' /C calc'!A0"
 Add CSV (Csv 'csv-cr-only.csv' ($header+"`r"+(Row)+"`r")) ''
 Add CSV (Csv 'csv-nul.csv' ($header+"`r`n"+(Row @{NAME=("A"+[char]0+"B")})+"`r`n")) ''
 Add CSV (Csv 'csv-empty.csv' '') 'no member rows'
 Add CSV (Csv 'csv-header-only.csv' ($header+"`r`n")) 'no member rows'
 Add CSV (Csv 'csv-missing-column.csv' ((@($keys[1..($keys.Count-1)]|ForEach-Object {'"'+$_+'"'}) -join ',')+"`r`n"+((Row) -replace '^"A",','')+"`r`n")) 'Missing RUAD column'
 Add CSV (Csv 'csv-duplicate-header.csv' ($header+',"NAME"'+"`r`n"+(Row)+',"A"'+"`r`n")) 'Repeated CSV header'
 Add CSV (Csv 'csv-extra-column.csv' ($header+',"INJECTED"'+"`r`n"+(Row)+',"A"'+"`r`n")) '35-column'
 Add CSV (Csv 'csv-short-row.csv' ($header+"`r`n"+((Row) -replace ',"A"$','')+"`r`n")) 'wrong column count'
 Add CSV (Csv 'csv-unclosed-quote.csv' ($header+"`r`n"+((Row) -replace '"$','')+"`r`n")) 'Unclosed CSV quote'
 Add CSV (Csv 'csv-mid-quote.csv' ($header+"`r`n"+((Row) -replace '^"A"','A"B')+"`r`n")) 'Unexpected CSV quote'
 Add CSV (Csv 'csv-after-quote.csv' ($header+"`r`n"+((Row) -replace '^"A"','"A"B')+"`r`n")) 'Unexpected text after CSV quote'
 Add CSV (Csv 'csv-long-cell.csv' ($header+"`r`n"+(Row @{NAME=('N'*4001)})+"`r`n")) 'exceeds 4,000'
 Add CSV (Csv 'csv-too-many-rows.csv' ($header+"`r`n"+(((1..10001)|ForEach-Object {Row})-join "`r`n")+"`r`n")) '10,000'
 Add CSV (Csv 'csv-too-large.csv' ($header+"`r`n"+(((1..1500)|ForEach-Object {Row @{NAME=('N'*3500)}})-join "`r`n")+"`r`n")) '5 MB'
 Add CSV (Save 'csv-utf16.csv' (Join-Bytes @([byte[]](0xFF,0xFE),[Text.Encoding]::Unicode.GetBytes($header+"`r`n"+(Row)+"`r`n")))) '*'
 Add CSV (Save 'csv-binary.csv' (Random-Bytes 8192)) '*'
 Add CSV (Save 'csv-invalid-utf8.csv' (Join-Bytes @($utf8.GetBytes($header+"`r`n"+'"'),[byte[]](0xC0,0xAF),$utf8.GetBytes('"'+((Row) -replace '^"A"','')+"`r`n")))) ''

 #--- NAVFIT98A --------------------------------------------------------------------
 $made=[string]$app.Eval('MakeNAVFITCases("'+$run+'")');if($made -notlike 'PASS:*'){throw $made}
 Add NAVFIT 'nav-control.accdb' ''
 Add NAVFIT (Save 'nav-empty.accdb' (New-Object byte[] 0)) '*'
 Add NAVFIT (Save 'nav-text.accdb' (Bytes 'This is not an Access database.')) '*'
 Add NAVFIT (Save 'nav-random.accdb' (Random-Bytes 65536)) '*'
 $valid=[IO.File]::ReadAllBytes((Join-Path $run 'nav-control.accdb'))
 Add NAVFIT (Save 'nav-truncated.accdb' ($valid[0..65535])) '*'
 Add NAVFIT 'nav-missing-table.accdb' '*'
 Add NAVFIT 'nav-missing-field.accdb' 'Missing NAVFIT field'
 Add NAVFIT 'nav-wrong-type.accdb' 'Unsupported NAVF'
 Add NAVFIT 'nav-extra-field.accdb' 'Unexpected NAVFIT fields'
 Add NAVFIT 'nav-password.accdb' '*'
 Add EXCHANGE 'exch-control.accdb' ''
 Add EXCHANGE 'exch-extra-table.accdb' '*'
 Add EXCHANGE 'exch-saved-query.accdb' 'saved queries'

 foreach($c in $cases){
  $path=Join-Path $run $c.File
  $call=switch($c.Kind){
   'PDF'{'MalformedPDF("'+$path+'","'+$c.Expect+'")'}
   'CSV'{'MalformedCSV("'+$path+'","'+$c.Expect+'","'+$c.Field+'","'+($c.Literal -replace '"','""')+'")'}
   'NAVFIT'{'MalformedNAVFIT("'+$path+'","'+$c.Expect+'",False)'}
   'EXCHANGE'{'MalformedNAVFIT("'+$path+'","'+$c.Expect+'",True)'}
  }
  $result=[string]$app.Eval($call)
  Write-Output ($c.Kind+' '+$c.File+': '+$result)
  if($result -like 'PASS:*'){$passed++}else{$failed+=($c.File+': '+$result)}
 }
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);Start-Sleep -Milliseconds 500;Remove-Item -LiteralPath $copy -ErrorAction SilentlyContinue}
if($failed.Count){throw ('Malformed-input cases failed: '+($failed -join ' | '))}
Write-Output ('PASS: '+$passed+' malformed-input and control cases (PDF, RUAD CSV, NAVFIT98A); rejected cleanly, no rows written, sources unchanged, state cleared, no temp files')
