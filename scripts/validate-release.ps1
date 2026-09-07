$ErrorActionPreference = "Stop"

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$distPath = Join-Path $projectRoot "dist"
$indexPath = Join-Path $distPath "index.html"

if (-not (Test-Path -LiteralPath $indexPath)) {
  throw "Build output is missing. Run npm run build before release validation."
}

$html = Get-Content -LiteralPath $indexPath -Raw -Encoding UTF8
$expectedCsp = "default-src 'none'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self'; img-src 'self' data:; connect-src https://arc.dbs.org;"
if (-not $html.Contains($expectedCsp)) {
  throw "The built task pane must retain the reviewed Content Security Policy."
}

if ($html.Contains('/@vite/client') -or $html.Contains('/src/main.ts') -or
    $html -notmatch '<link[^>]+rel="stylesheet"[^>]+href="/assets/[^" ]+\.css"') {
  throw "Host validation must serve built scripts and an external stylesheet, not Vite's CSP-blocked development client."
}

$package = Get-Content -LiteralPath (Join-Path $projectRoot "package.json") -Raw | ConvertFrom-Json
if (-not $package.scripts.'desktop:start'.StartsWith('npm run build && ') -or
    -not $package.scripts.'desktop:start'.Contains('--dev-server "npm run desktop:serve"') -or
    $package.scripts.'desktop:serve' -ne 'vite preview --host localhost --port 3000 --strictPort') {
  throw "Desktop sideload must build and serve the exact production assets."
}

$localhostHits = Get-ChildItem -LiteralPath $distPath -Recurse -File |
  Select-String -SimpleMatch "https://localhost:3000" -ErrorAction Stop
if ($localhostHits) {
  throw "Production web assets must not contain the development origin."
}

$fixtureManifest = Join-Path $distPath "release-validation-manifest.xml"
$builder = Join-Path $PSScriptRoot "build-production-manifest.ps1"

function Assert-BuilderRejects {
  param(
    [scriptblock]$Action,
    [string]$Label
  )

  $rejected = $false
  try {
    & $Action | Out-Null
  }
  catch {
    $rejected = $true
  }

  if (-not $rejected) {
    throw "Production manifest builder accepted $Label."
  }
}

Assert-BuilderRejects -Label "a localhost production origin" -Action {
  & $builder `
    -PublicBaseUrl "https://localhost:4443/verseform-word" `
    -SupportUrl "https://support.example.invalid/verseform-word" `
    -ProviderName "Release validation fixture" `
    -OutputPath $fixtureManifest
}

Assert-BuilderRejects -Label "an insecure support URL" -Action {
  & $builder `
    -PublicBaseUrl "https://distribution.example.invalid/verseform-word" `
    -SupportUrl "http://support.example.invalid/verseform-word" `
    -ProviderName "Release validation fixture" `
    -OutputPath $fixtureManifest
}

& $builder `
  -PublicBaseUrl "https://distribution.example.invalid/verseform-word" `
  -SupportUrl "https://support.example.invalid/verseform-word" `
  -ProviderName "Release validation fixture" `
  -OutputPath $fixtureManifest

Write-Output "Built assets and production-manifest transformation passed"
