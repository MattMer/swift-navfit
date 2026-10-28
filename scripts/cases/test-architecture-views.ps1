param([Parameter(Mandatory=$true)][string]$BuildPath)
# Fails when a generated architecture view (docs/architecture/views) no longer
# matches tracked source and the build, or an OV-6b state list drifts from the schema.
& (Join-Path (Split-Path $PSScriptRoot -Parent) 'generate-architecture-views.ps1') -BuildPath $BuildPath -Check
