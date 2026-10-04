param([string]$OutputPath=(Join-Path $PSScriptRoot '..\output\ux\Swift_NAVFIT_UI_Prototype.html'))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$prototype=Join-Path $root 'prototype'
$html=[IO.File]::ReadAllText((Join-Path $prototype 'index.html'))
$css=[IO.File]::ReadAllText((Join-Path $prototype 'styles.css'))
$parts=@()
foreach($name in @('fields.mjs','traits.mjs','model.mjs','app.mjs')){
    $part=[IO.File]::ReadAllText((Join-Path $prototype $name))
    $part=[regex]::Replace($part,'(?m)^import .*?;\r?\n','')
    $part=[regex]::Replace($part,'(?m)^export ','')
    $parts+=$part
}
$html=$html.Replace('<link rel="stylesheet" href="styles.css">',('<style>'+$css+'</style>'))
$html=$html.Replace('<script type="module" src="app.mjs"></script>',('<script type="module">'+($parts -join "`n")+'</script>'))
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
[void](New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath -Parent))
[IO.File]::WriteAllText($OutputPath,$html,[Text.UTF8Encoding]::new($false))
[pscustomobject]@{Output=$OutputPath;SHA256=(Get-FileHash -LiteralPath $OutputPath).Hash;Status='Built self-contained design prototype; not an Access application'}|ConvertTo-Json
