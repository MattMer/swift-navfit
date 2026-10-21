$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'vba-source-normalization.ps1')
$checks=0
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}
Check ((NormalizeSwiftVBA 'Call RefreshLibrary') -ceq (NormalizeSwiftVBA 'Call RefreshLibrary()')) 'Access empty Call parentheses differ'
Check ((NormalizeSwiftVBA 'CALL RefreshLibrary( )') -ceq (NormalizeSwiftVBA 'Call RefreshLibrary')) 'Call casing/spacing differ'
Check ((NormalizeSwiftVBA 'x = "Call RefreshLibrary()"') -cne (NormalizeSwiftVBA 'x = "Call RefreshLibrary"')) 'Quoted Call text was normalized'
Check ((NormalizeSwiftVBA 'x = "A"') -cne (NormalizeSwiftVBA 'x = "a"')) 'Quoted casing was normalized'
Check ((NormalizeSwiftVBA 'x = "A "') -cne (NormalizeSwiftVBA 'x = "A"')) 'Quoted trailing space was lost'
Check ((NormalizeSwiftVBA 'Call Foo(1)') -cne (NormalizeSwiftVBA 'Call Foo(2)')) 'Call argument difference was lost'
Check ((NormalizeSwiftVBA 'x = "can''t"') -cne (NormalizeSwiftVBA 'x = "cant"')) 'Quoted apostrophe was lost'
Check ((NormalizeSwiftVBA "x=1 ' comment") -ceq (NormalizeSwiftVBA 'X = 1')) 'Comments or code casing differ'
Write-Output ('PASS: '+$checks+' VBA source-normalization fidelity checks')
