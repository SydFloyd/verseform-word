[CmdletBinding()]
param(
  [string]$ManifestPath,

  [ValidateSet("Development", "Production")]
  [string]$Mode = "Development",

  [string]$ExpectedBaseUrl
)

$ErrorActionPreference = "Stop"

if (-not $ManifestPath) {
  $ManifestPath = Join-Path $PSScriptRoot "..\manifest.xml"
}

$resolvedManifestPath = [IO.Path]::GetFullPath($ManifestPath)
[xml]$manifest = Get-Content -LiteralPath $resolvedManifestPath -Raw -Encoding UTF8

$root = $manifest.DocumentElement
if ($null -eq $root -or $root.LocalName -ne "OfficeApp") {
  throw "manifest.xml must contain an OfficeApp root element."
}

$requiredText = @{
  "Id" = "82d488fc-c5a7-4aa5-8133-d56c32ef9a69"
  "Version" = "1.0.0.5"
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

$providerName = $root.SelectSingleNode("./*[local-name()='ProviderName']")
if ($null -eq $providerName -or -not $providerName.InnerText.Trim() -or $providerName.InnerText.Length -gt 125) {
  throw "manifest.xml must expose a provider name between 1 and 125 characters."
}

$support = $root.SelectSingleNode("./*[local-name()='SupportUrl']")
if ($null -eq $support -or -not $support.GetAttribute("DefaultValue").StartsWith("https://")) {
  throw "manifest.xml must expose an HTTPS support URL."
}

$manifestHost = $root.SelectSingleNode("./*[local-name()='Hosts']/*[local-name()='Host' and @Name='Document']")
if ($null -eq $manifestHost) {
  throw "manifest.xml must target the Word Document host."
}

$source = $root.SelectSingleNode("./*[local-name()='DefaultSettings']/*[local-name()='SourceLocation']")
if ($null -eq $source) {
  throw "manifest.xml must expose a source location."
}

if ($Mode -eq "Development" -and -not $source.GetAttribute("DefaultValue").StartsWith("https://localhost:3000/")) {
  throw "The development manifest must use the trusted HTTPS localhost origin."
}

$wordSets = @($root.SelectNodes(".//*[local-name()='Set' and @Name='WordApi']"))
if ($wordSets.Count -lt 2) {
  throw "manifest.xml must require WordApi in both base and command surfaces."
}

foreach ($set in $wordSets) {
  $minimum = $set.GetAttribute("MinVersion")
  if (-not $minimum) {
    $minimum = $set.ParentNode.GetAttribute("DefaultMinVersion")
  }
  if ($minimum -ne "1.7") {
    throw "manifest.xml must require WordApi 1.7."
  }
}

$sharedRuntimeSets = @($root.SelectNodes(".//*[local-name()='Set' and @Name='SharedRuntime']"))
if ($sharedRuntimeSets.Count -lt 2) {
  throw "manifest.xml must require SharedRuntime in both base and command surfaces."
}
foreach ($set in $sharedRuntimeSets) {
  $minimum = $set.GetAttribute("MinVersion")
  if (-not $minimum) {
    $minimum = $set.ParentNode.GetAttribute("DefaultMinVersion")
  }
  if ($minimum -ne "1.1") {
    throw "manifest.xml must require SharedRuntime 1.1."
  }
}

$runtimes = @($root.SelectNodes(".//*[local-name()='Runtimes']/*[local-name()='Runtime']"))
if ($runtimes.Count -ne 1 -or $runtimes[0].GetAttribute("resid") -ne "Taskpane.Url" -or $runtimes[0].GetAttribute("lifetime") -ne "long") {
  throw "manifest.xml must use the task-pane page as one long shared runtime."
}

$functionFile = $root.SelectSingleNode(".//*[local-name()='FunctionFile']")
if ($null -eq $functionFile -or $functionFile.GetAttribute("resid") -ne "Taskpane.Url") {
  throw "manifest.xml commands must share the task-pane runtime."
}

$enableAction = $root.SelectSingleNode(".//*[local-name()='Control' and @id='Verseform.Enable']/*[local-name()='Action' and @*[local-name()='type']='ExecuteFunction']")
if ($null -eq $enableAction -or $enableAction.SelectSingleNode("./*[local-name()='FunctionName']").InnerText -ne "enableVerseform") {
  throw "manifest.xml must expose the background Enable Verseform command."
}

$fillActions = @($root.SelectNodes(".//*[local-name()='Control' and (@id='Verseform.Fill' or @id='Verseform.FillContext')]/*[local-name()='Action' and @*[local-name()='type']='ExecuteFunction']"))
if ($fillActions.Count -ne 2 -or @($fillActions | Where-Object { $_.SelectSingleNode("./*[local-name()='FunctionName']").InnerText -ne "fillScripture" }).Count -ne 0) {
  throw "manifest.xml must expose Fill Scripture on the ribbon and text context menu."
}
$contextMenu = $root.SelectSingleNode(".//*[local-name()='ExtensionPoint' and @*[local-name()='type']='ContextMenu']/*[local-name()='OfficeMenu' and @id='ContextMenuText']")
if ($null -eq $contextMenu) {
  throw "manifest.xml must expose Fill Scripture in Word's selected-text context menu."
}

$paneAction = $root.SelectSingleNode(".//*[local-name()='Control' and @id='Verseform.ShowTaskpane']/*[local-name()='Action' and @*[local-name()='type']='ShowTaskpane']")
if ($null -eq $paneAction) {
  throw "manifest.xml must expose the optional preview and settings pane."
}
$taskpaneIds = @($root.SelectNodes(".//*[local-name()='Action']/*[local-name()='TaskpaneId']"))
if ($taskpaneIds.Count -ne 0) {
  throw "Microsoft's long shared-runtime configuration must not declare TaskpaneId."
}
$paneSource = $paneAction.SelectSingleNode("./*[local-name()='SourceLocation']")
if ($null -eq $paneSource -or $paneSource.GetAttribute("resid") -ne "Taskpane.Url") {
  throw "The optional pane must use the same resource ID as the long runtime and function file."
}

if ($null -ne $root.SelectSingleNode("./*[local-name()='ExtendedOverrides']")) {
  throw "Do not advertise an add-in shortcut until its host registration is proved."
}

$urls = @($root.SelectNodes(".//*[(local-name()='SourceLocation' or local-name()='Url' or local-name()='Image' or local-name()='IconUrl' or local-name()='HighResolutionIconUrl' or local-name()='SupportUrl') and @DefaultValue]"))
foreach ($node in $urls) {
  $value = $node.GetAttribute("DefaultValue")
  if (-not $value -or -not $value.StartsWith("https://")) {
    throw "Every manifest resource URL must use HTTPS."
  }
}

if ($Mode -eq "Production") {
  if (-not $ExpectedBaseUrl) {
    throw "ExpectedBaseUrl is required in Production mode."
  }

  try {
    $expectedUri = [Uri]$ExpectedBaseUrl
  }
  catch {
    throw "ExpectedBaseUrl must be an absolute HTTPS URL."
  }
  if (-not $expectedUri.IsAbsoluteUri -or $expectedUri.Scheme -ne "https" -or $expectedUri.UserInfo -or $expectedUri.Query -or $expectedUri.Fragment) {
    throw "ExpectedBaseUrl must be a credential-free HTTPS URL without a query or fragment."
  }

  $normalizedBase = $expectedUri.AbsoluteUri.TrimEnd("/")
  if ($expectedUri.IsLoopback) {
    throw "A production manifest cannot use localhost."
  }

  try {
    $supportUri = [Uri]$support.GetAttribute("DefaultValue")
  }
  catch {
    throw "The production support URL must be an absolute HTTPS URL."
  }
  if (-not $supportUri.IsAbsoluteUri -or $supportUri.Scheme -ne "https" -or $supportUri.IsLoopback -or $supportUri.UserInfo -or $supportUri.Query -or $supportUri.Fragment) {
    throw "The production support URL must be credential-free public HTTPS without a query or fragment."
  }

  $appResources = @(
    $root.SelectNodes("./*[local-name()='IconUrl' or local-name()='HighResolutionIconUrl']")
    $root.SelectNodes("./*[local-name()='DefaultSettings']/*[local-name()='SourceLocation']")
    $root.SelectNodes(".//*[local-name()='Image' and @DefaultValue]")
    $root.SelectNodes(".//*[local-name()='Url' and (@id='Commands.Url' or @id='Taskpane.Url')]")
  )

  foreach ($nodeGroup in $appResources) {
    foreach ($node in @($nodeGroup)) {
      $value = $node.GetAttribute("DefaultValue")
      if (-not $value.StartsWith("$normalizedBase/", [StringComparison]::OrdinalIgnoreCase)) {
        throw "Every production app resource must use the single approved public base URL."
      }
    }
  }

  foreach ($node in $urls) {
    if ($node.GetAttribute("DefaultValue") -match "^https://localhost(?::|/|$)") {
      throw "A production manifest cannot contain a localhost URL."
    }
  }
}

Write-Output "$Mode manifest structural contract passed: $resolvedManifestPath"
