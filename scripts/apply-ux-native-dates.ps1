if(-not $script:swiftManagedBuild){foreach($m in @('modNAVFITAdapter','modUXNativeGroups','modUXNAVFIT')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
