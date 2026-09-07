[CmdletBinding()]
param(
  [string]$ManifestPath
)

$ErrorActionPreference = "Stop"

if (-not $ManifestPath) {
  $ManifestPath = Join-Path $PSScriptRoot "..\dist\verseform-for-word-manifest.xml"
}

$fullManifestPath = [IO.Path]::GetFullPath($ManifestPath)
if (-not (Test-Path -LiteralPath $fullManifestPath -PathType Leaf)) {
  throw "Production manifest not found at $fullManifestPath. Build it with npm run manifest:production first."
}

$manifestText = Get-Content -LiteralPath $fullManifestPath -Raw -Encoding UTF8
if ($manifestText.Contains("localhost") -or $manifestText.Contains(".invalid")) {
  throw "Marketplace validation refuses development or fixture manifest URLs."
}

& npx.cmd --yes office-addin-manifest@2.1.6 validate -p $fullManifestPath
if ($LASTEXITCODE -ne 0) {
  throw "Microsoft Marketplace production validation failed for $fullManifestPath."
}

Write-Output "Microsoft Marketplace production validation passed: $fullManifestPath"
