# Embed the already-supported, tracked native blank; export never asks a user to
# supply implementation resources. Small procedures stay below VBA size limits.
$nativeTemplate=Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb'
if((Get-FileHash -LiteralPath $nativeTemplate).Hash -cne 'DD3806E2DAAAE54B3F9B4EC07D59FB0654B46AADECEDCFF54F57DCC75167B2FA'){throw 'Review the changed native blank before embedding it'}
$encoded=[Convert]::ToBase64String([IO.File]::ReadAllBytes($nativeTemplate))
$templateCode=[Collections.Generic.List[string]]::new();$templateCode.Add('Attribute VB_Name = "modNAVFITTemplate"');$templateCode.Add('Option Explicit')
$templateCode.Add('Public Function NAVFITTemplateHash() As String');$templateCode.Add(' NAVFITTemplateHash = "'+(Get-FileHash -LiteralPath $nativeTemplate).Hash+'"');$templateCode.Add('End Function')
$chunks=[int][math]::Ceiling($encoded.Length/16384)
$templateCode.Add('Public Function NAVFITTemplateData() As String');$templateCode.Add(' Dim result As String')
for($i=0;$i -lt $chunks;$i++){$templateCode.Add(' result = result & NativeTemplateChunk'+$i+'()')}
$templateCode.Add(' NAVFITTemplateData = result');$templateCode.Add('End Function')
for($i=0;$i -lt $chunks;$i++){
 $part=$encoded.Substring($i*16384,[math]::Min(16384,$encoded.Length-$i*16384))
 $templateCode.Add('Private Function NativeTemplateChunk'+$i+'() As String');$templateCode.Add(' Dim result As String')
 for($offset=0;$offset -lt $part.Length;$offset+=512){$templateCode.Add(' result = result & "'+$part.Substring($offset,[math]::Min(512,$part.Length-$offset))+'"')}
 $templateCode.Add(' NativeTemplateChunk'+$i+' = result');$templateCode.Add('End Function')
}
$templateModule=Join-Path $root 'build/modNAVFITTemplate.bas';[IO.File]::WriteAllLines($templateModule,$templateCode,[Text.UTF8Encoding]::new($false));$app.LoadFromText(5,'modNAVFITTemplate',$templateModule)
