param([string]$InstructionPath=(Join-Path $PSScriptRoot '../BUPERSINST 1610.10.pdf'),[string]$InstructionTextPath=(Join-Path $PSScriptRoot '../planning/evidence/BUPERSINST 1610.10.txt'))
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$sources=@((Get-Content (Join-Path $root 'src/modLiveValidation.bas') -Raw),(Get-Content (Join-Path $root 'src/modValidationReferences.bas') -Raw))
$definitions=Get-Content (Join-Path $root 'src/modValidation.bas') -Raw
$expected=[regex]::Match($definitions,'NAVY_RULE_SOURCE_HASH As String = "([A-F0-9]+)"').Groups[1].Value
if((Get-FileHash -LiteralPath $InstructionPath).Hash -ne $expected){throw 'Instruction differs from the rule source fingerprint; reverify citations before updating the profile.'}
if((Get-FileHash -LiteralPath $InstructionTextPath).Hash -ne '318926467046A731A0D77A5F57B6FE2849A3C886C2CF9287579B55824FF11F1F'){throw 'Recorded PDF extraction differs; reverify source text and locators before changing this fingerprint.'}
function NormalizeQuote([string]$value){return (($value.Replace([string][char]0x201c,'"').Replace([string][char]0x201d,'"') -replace '\s+',' ').Trim())}
$pageText=Get-Content -LiteralPath $InstructionTextPath -Raw -Encoding utf8
$policy=NormalizeQuote (Get-Content -LiteralPath $InstructionTextPath -Raw -Encoding utf8)
$count=0
foreach($source in $sources){
foreach($match in [regex]::Matches($source,'excerpt = "((?:[^"\r\n]|"")*)"')){
$excerpt=NormalizeQuote ($match.Groups[1].Value.Replace('""','"'))
if(-not $policy.Contains($excerpt)){throw ('Excerpt differs from supplied instruction: '+$excerpt)}
$prefix=$source.Substring(0,$match.Index)
$pageMatches=[regex]::Matches($prefix,'pdfPage = ([0-9]+)')
$page=[int]$pageMatches[$pageMatches.Count-1].Groups[1].Value
$pageBody=[regex]::Match($pageText,'(?s)--- PDF PAGE '+$page+' ---\s*(.*?)(?=--- PDF PAGE|$)').Groups[1].Value
if(-not (NormalizeQuote $pageBody).Contains($excerpt)){throw ('Excerpt is not on claimed PDF page '+$page+': '+$excerpt)}
$count++
}
}
if($count -lt 20){throw 'Expected citation catalogue is incomplete'}
Write-Output ('PASS: '+$count+' policy excerpts and page locators match fingerprinted PDF extraction (whitespace/quotation typography normalized)')
