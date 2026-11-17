param(
    [Parameter(Mandatory=$true)][string]$OutputFolder,
    [Parameter(Mandatory=$true)][string]$Version
)
# One download with everything: Swift_NAVFIT_<version>.zip holds the application,
# the Demo pair and the user guide. package-release.ps1 calls this after it checks
# the release files; it can also be run on an existing publish folder.
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$OutputFolder=[IO.Path]::GetFullPath($OutputFolder)
$files=[ordered]@{}
foreach($name in ('Swift_NAVFIT_'+$Version+'.accdb'),('Swift_NAVFIT_Demo_'+$Version+'.accdb'),('Swift_NAVFIT_Demo_data_'+$Version+'.accdb')){$files[$name]=Join-Path $OutputFolder $name}
$files['Swift_NAVFIT_User_Guide.pdf']=Join-Path $root 'docs\Swift_NAVFIT_User_Guide.pdf'
foreach($path in $files.Values){if(-not(Test-Path -LiteralPath $path)){throw ('Missing release file: '+$path)}}
$zip=Join-Path $OutputFolder ('Swift_NAVFIT_'+$Version+'.zip')
if(Test-Path -LiteralPath $zip){throw ('Refusing to overwrite '+(Split-Path $zip -Leaf))}

Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
$archive=[IO.Compression.ZipFile]::Open($zip,[IO.Compression.ZipArchiveMode]::Create)
try{foreach($name in $files.Keys){[void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$files[$name],$name,[IO.Compression.CompressionLevel]::Optimal)}}
finally{$archive.Dispose()}

#Every entry must read back identical to the file it came from.
$sha=[Security.Cryptography.SHA256]::Create()
$archive=[IO.Compression.ZipFile]::OpenRead($zip)
try{
    if($archive.Entries.Count -ne $files.Count){throw 'The zip has the wrong number of files'}
    foreach($entry in $archive.Entries){
        if(-not $files.Contains($entry.FullName)){throw ('Unexpected file in the zip: '+$entry.FullName)}
        $stream=$entry.Open();try{$hash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}finally{$stream.Dispose()}
        if($hash -ne (Get-FileHash -LiteralPath $files[$entry.FullName]).Hash){throw ('Zip entry differs from its source: '+$entry.FullName)}
    }
}finally{$archive.Dispose();$sha.Dispose()}
Write-Output ((Split-Path $zip -Leaf)+'  '+(Get-Item -LiteralPath $zip).Length+' bytes  SHA256 '+(Get-FileHash -LiteralPath $zip).Hash)
Write-Output ('PASS: zip holds the application, the Demo pair and the user guide, each identical to its source')
