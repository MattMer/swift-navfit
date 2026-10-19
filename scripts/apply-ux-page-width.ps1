foreach($m in @('modReportPages','modUXSurface')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
