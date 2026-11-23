[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [uri] $ItemUrl,

  [Parameter(Mandatory)]
  [string] $ActionItemInstanceId,

  [switch] $Launch
)

$ErrorActionPreference = "Stop"

if (
  $ItemUrl.Scheme -ne "https" -or
  $ItemUrl.Host -ne "vnext.s360.msftcloudes.com" -or
  $ItemUrl.AbsolutePath.TrimEnd("/") -ne "/blades/security"
) {
  throw "Expected an S360 security item URL under https://vnext.s360.msftcloudes.com/blades/security."
}

$decodedQuery = [uri]::UnescapeDataString($ItemUrl.Query.TrimStart("?"))
$kpiMatch = [regex]::Match(
  $decodedQuery,
  "(?:^|[&~])KPI:(?<id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})(?:[&~]|$)",
  [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)
if (-not $kpiMatch.Success) {
  throw "The S360 security item URL does not contain a valid KPI ID."
}
$kpiId = ([guid] $kpiMatch.Groups["id"].Value).ToString("D")

if ($ActionItemInstanceId -notmatch "^[A-Za-z0-9][A-Za-z0-9._:|/-]{0,255}$") {
  throw "The KpiActionItemId must be a single identifier containing only letters, numbers, period, underscore, colon, pipe, slash, or hyphen."
}

$repoRootOutput = & git rev-parse --show-toplevel 2>$null
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($repoRootOutput)) {
  throw "Run S360 remediation from the target Git repository."
}
$repoRoot = [System.IO.Path]::GetFullPath(($repoRootOutput | Select-Object -First 1))

$originRemoteOutput = & git -C $repoRoot remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($originRemoteOutput)) {
  throw "The current Git repository does not have an origin remote to compare with the S360 action item."
}
$originRemote = [string] ($originRemoteOutput | Select-Object -First 1)

$sessionId = [guid]::NewGuid().ToString("D")
$prompt = @"
You are being launched from the Service 360 (S360) portal to resolve a specific S360 action item. This session already includes the service360-breeze-official plugin and the s360-breeze MCP server — do not install anything or suggest starting a new session. Help the user resolve the S360 action item for KPI $kpiId (action item instance id / KpiActionItemId: $ActionItemInstanceId). Use the 'Production' environment for any tools that accept an optional 'environment' input for this session.

Before editing, identify the repository associated with the returned action item and compare it with the current repository origin '$originRemote'. Normalize HTTPS and SSH Git remote forms when comparing. If they do not match, stop without modifying files and report the mismatch.

If the item exists and the repository matches, investigate its requirements, implement the smallest complete fix in the current repository, and run focused validation. Preserve existing workspace changes. Report the KPI ID, row-level action-item instance ID, root cause, files changed, and validation results. The source S360 link is: $($ItemUrl.ToString())
"@.Trim()

$arguments = @(
  "copilot",
  "--plugin",
  "mp:service360-breeze-official@company",
  "--agent",
  "service360-breeze-official:s360-action-item-resolver",
  "--mcp",
  "s360-remediation: s360-breeze",
  "--session-id",
  $sessionId,
  "--prompt",
  $prompt,
  "--autopilot",
  "--assisted-approval",
  "--no-ask-user",
  "--disable-mcp-server",
  "s360-breeze"
)

$context = [ordered]@{
  itemUrl = $ItemUrl.ToString()
  kpiId = $kpiId
  actionItemInstanceId = $ActionItemInstanceId
  sessionId = $sessionId
  repositoryRoot = $repoRoot
  originRemote = $originRemote
  executable = "agency"
  arguments = $arguments
}

if (-not $Launch) {
  $context | ConvertTo-Json -Depth 10
  return
}

$agencyCommand = Get-Command agency -ErrorAction SilentlyContinue
if ($null -eq $agencyCommand) {
  throw "Agency is not installed or is not available on PATH."
}

Push-Location $repoRoot
try {
  & $agencyCommand.Source @arguments
  if ($LASTEXITCODE -ne 0) {
    throw "The S360 Agency remediation session exited with code $LASTEXITCODE."
  }
}
finally {
  Pop-Location
}
