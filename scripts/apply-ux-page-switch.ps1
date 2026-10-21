if(-not $script:swiftManagedBuild){foreach($m in @('modReportPages','modUXVisual','modUXHeaderReview')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
