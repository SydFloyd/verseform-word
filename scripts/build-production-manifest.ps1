[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$PublicBaseUrl,

  [Parameter(Mandatory = $true)]
  [string]$SupportUrl,

  [Parameter(Mandatory = $true)]
  [string]$ProviderName,

  [string]$OutputPath
)

$ErrorActionPreference = "Stop"

if (-not $OutputPath) {
  $OutputPath = Join-Path $PSScriptRoot "..\dist\verseform-for-word-manifest.xml"
}

function Get-StrictHttpsUri {
  param(
    [string]$Value,
    [string]$Label
  )

  try {
    $uri = [Uri]$Value
  }
  catch {
    throw "$Label must be an absolute HTTPS URL."
  }

  if (-not $uri.IsAbsoluteUri -or $uri.Scheme -ne "https") {
    throw "$Label must be an absolute HTTPS URL."
  }
  if ($uri.IsLoopback) {
    throw "$Label must not use a loopback host for a production package."
  }
  if ($uri.UserInfo -or $uri.Query -or $uri.Fragment) {
    throw "$Label must not contain credentials, a query, or a fragment."
  }

  return $uri
}

$publicUri = Get-StrictHttpsUri -Value $PublicBaseUrl -Label "PublicBaseUrl"
$supportUri = Get-StrictHttpsUri -Value $SupportUrl -Label "SupportUrl"
$normalizedBase = $publicUri.AbsoluteUri.TrimEnd("/")
$normalizedSupport = $supportUri.AbsoluteUri.TrimEnd("/")
$trimmedProvider = $ProviderName.Trim()

if (-not $trimmedProvider -or $trimmedProvider.Length -gt 125) {
  throw "ProviderName must contain between 1 and 125 characters."
}

$sourcePath = Join-Path $PSScriptRoot "..\manifest.xml"
[xml]$manifest = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
$root = $manifest.DocumentElement

$providerNode = $root.SelectSingleNode("./*[local-name()='ProviderName']")
$providerNode.InnerText = $trimmedProvider

$supportNode = $root.SelectSingleNode("./*[local-name()='SupportUrl']")
$supportNode.SetAttribute("DefaultValue", $normalizedSupport)

$learnMoreNode = $root.SelectSingleNode(".//*[local-name()='Url' and @id='GetStarted.LearnMoreUrl']")
$learnMoreNode.SetAttribute("DefaultValue", $normalizedSupport)

$developmentOrigin = "https://localhost:3000"
$replacedResources = 0
foreach ($node in @($root.SelectNodes(".//*[@DefaultValue]"))) {
  $value = $node.GetAttribute("DefaultValue")
  if ($value.StartsWith($developmentOrigin, [StringComparison]::OrdinalIgnoreCase)) {
    $relative = $value.Substring($developmentOrigin.Length)
    $node.SetAttribute("DefaultValue", "$normalizedBase$relative")
    $replacedResources += 1
  }
}

if ($replacedResources -ne 8) {
  throw "Expected to replace 8 development resource URLs, replaced $replacedResources."
}

$fullOutputPath = [IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $fullOutputPath
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$settings = [Xml.XmlWriterSettings]::new()
$settings.Encoding = [Text.UTF8Encoding]::new($false)
$settings.Indent = $true
$writer = [Xml.XmlWriter]::Create($fullOutputPath, $settings)
try {
  $manifest.Save($writer)
}
finally {
  $writer.Dispose()
}

$validator = Join-Path $PSScriptRoot "validate-manifest.ps1"
& $validator -ManifestPath $fullOutputPath -Mode Production -ExpectedBaseUrl $normalizedBase
Write-Output "Production manifest written to $fullOutputPath"
