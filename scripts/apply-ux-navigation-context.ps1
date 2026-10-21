# Explicit group defaults context and a single comparison visibility pass.
foreach($m in @('modUX','modUXBlockTools','modUXGroupPages','modUXFeedback','modUXDashboard')){
if(-not $script:swiftManagedBuild){ $app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
}
