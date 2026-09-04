$ErrorActionPreference = "Stop"

$manifestPath = Join-Path $PSScriptRoot "..\manifest.xml"
[xml]$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8

$root = $manifest.DocumentElement
if ($null -eq $root -or $root.LocalName -ne "OfficeApp") {
  throw "manifest.xml must contain an OfficeApp root element."
}

$requiredText = @{
  "Id" = "82d488fc-c5a7-4aa5-8133-d56c32ef9a69"
  "Version" = "0.1.0.0"
  "Permissions" = "ReadWriteDocument"
}

foreach ($entry in $requiredText.GetEnumerator()) {
  $node = $root.SelectSingleNode("./*[local-name()='$($entry.Key)']")
  if ($null -eq $node -or $node.InnerText -ne $entry.Value) {
    throw "manifest.xml has an invalid $($entry.Key)."
  }
}

$displayName = $root.SelectSingleNode("./*[local-name()='DisplayName']")
if ($null -eq $displayName -or $displayName.GetAttribute("DefaultValue") -ne "Verseform for Word") {
  throw "manifest.xml must expose the approved working name."
}

$manifestHost = $root.SelectSingleNode("./*[local-name()='Hosts']/*[local-name()='Host' and @Name='Document']")
if ($null -eq $manifestHost) {
  throw "manifest.xml must target the Word Document host."
}

$source = $root.SelectSingleNode("./*[local-name()='DefaultSettings']/*[local-name()='SourceLocation']")
if ($null -eq $source -or -not $source.GetAttribute("DefaultValue").StartsWith("https://localhost:3000/")) {
  throw "manifest.xml must use the trusted HTTPS development origin."
}

$wordSets = @($root.SelectNodes(".//*[local-name()='Set' and @Name='WordApi']"))
if ($wordSets.Count -lt 2) {
  throw "manifest.xml must require WordApi in both base and command surfaces."
}

foreach ($set in $wordSets) {
  $minimum = $set.ParentNode.GetAttribute("DefaultMinVersion")
  if ($minimum -ne "1.7") {
    throw "manifest.xml must require WordApi 1.7."
  }
}

$urls = @($root.SelectNodes(".//*[local-name()='SourceLocation' or local-name()='Url' or local-name()='IconUrl' or local-name()='HighResolutionIconUrl']"))
foreach ($node in $urls) {
  $value = $node.GetAttribute("DefaultValue")
  if ($value -and -not $value.StartsWith("https://")) {
    throw "Every manifest resource URL must use HTTPS."
  }
}

Write-Output "manifest.xml structural contract passed"
