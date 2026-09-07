$ErrorActionPreference = "Stop"

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$distPath = Join-Path $projectRoot "dist"
$indexPath = Join-Path $distPath "index.html"
$validator = Join-Path $PSScriptRoot "validate-manifest.ps1"

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

$pilotBaseUrl = "https://verseform-word.kmproto.com"
$publicPilotManifest = Join-Path $projectRoot "public\verseform-for-word.xml"
$builtPilotManifest = Join-Path $distPath "verseform-for-word.xml"
$builtInstallPage = Join-Path $distPath "install.html"
$vercelConfigPath = Join-Path $projectRoot "vercel.json"

if (-not (Test-Path -LiteralPath $publicPilotManifest -PathType Leaf) -or
    -not (Test-Path -LiteralPath $builtPilotManifest -PathType Leaf)) {
  throw "The friend-pilot manifest must exist in public and the built output."
}

& $validator -ManifestPath $publicPilotManifest -Mode Production -ExpectedBaseUrl $pilotBaseUrl
$publicPilotBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($publicPilotManifest))
$builtPilotBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($builtPilotManifest))
if ($publicPilotBytes -ne $builtPilotBytes) {
  throw "Vite did not copy the reviewed friend-pilot manifest exactly."
}

if (-not (Test-Path -LiteralPath $builtInstallPage -PathType Leaf)) {
  throw "The friend-pilot install page is missing from the built output."
}
$installHtml = Get-Content -LiteralPath $builtInstallPage -Raw -Encoding UTF8
if (-not $installHtml.Contains('href="/verseform-for-word.xml"') -or
    -not $installHtml.Contains("Upload My Add-in") -or
    -not $installHtml.Contains("This manual web pilot does not install Verseform into Word Desktop.")) {
  throw "The friend-pilot install page must link the manifest and state its exact persistence boundary."
}

if (-not (Test-Path -LiteralPath $vercelConfigPath -PathType Leaf)) {
  throw "The reviewed Vercel hosting contract is missing."
}
$vercelText = Get-Content -LiteralPath $vercelConfigPath -Raw -Encoding UTF8
$vercel = $vercelText | ConvertFrom-Json
if ($vercel.buildCommand -ne "npm run build" -or $vercel.outputDirectory -ne "dist" -or
    -not $vercelText.Contains("Content-Disposition") -or
    $vercelText.Contains("X-Frame-Options") -or $vercelText.Contains("frame-ancestors 'none'")) {
  throw "Vercel must deploy the reviewed build, download the manifest, and permit Office framing."
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
