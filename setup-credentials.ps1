$configDir = Join-Path $env:USERPROFILE ".config\daisys-dough"
New-Item -ItemType Directory -Force -Path $configDir | Out-Null

$download = Get-ChildItem -Path (Join-Path $env:USERPROFILE "Downloads\client_secret*.json") -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $download) {
    Write-Error "No client_secret*.json found in Downloads. Download Desktop OAuth JSON from Google Cloud Console first."
    exit 1
}

$dest = Join-Path $configDir "oauth_client.json"
Copy-Item -Path $download.FullName -Destination $dest -Force

$projectToken = Join-Path $PSScriptRoot "token.json"
$configToken = Join-Path $configDir "token.json"
if ((Test-Path $projectToken) -and -not (Test-Path $configToken)) {
    Copy-Item -Path $projectToken -Destination $configToken -Force
    Write-Host "Moved token.json to $configToken"
}

Write-Host "OAuth client installed:"
Write-Host "  $dest"
Write-Host ""
Write-Host "Update .env if needed:"
Write-Host "  GOOGLE_OAUTH_CLIENT_PATH=$dest"
Write-Host "  GOOGLE_TOKEN_PATH=$configToken"
