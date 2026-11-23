$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "start-s360-remediation.ps1"
$kpiId = "c9205e25-ee12-4aa1-8c1f-693df581d500"
$actionItemInstanceId = "000109caddb614e6dc629fc3b2d00dd6e03c1e452ae888404a0a85dd679f4341"
$itemUrl = "https://vnext.s360.msftcloudes.com/blades/security?blade=AssignedTo:All~KPI:$kpiId~Tab:Summary"

function Assert-Rejected {
  param(
    [Parameter(Mandatory)][string] $Name,
    [Parameter(Mandatory)][uri] $Url,
    [Parameter(Mandatory)][string] $InstanceId,
    [Parameter(Mandatory)][string] $ExpectedMessage
  )

  try {
    & $scriptPath `
      -ItemUrl $Url `
      -ActionItemInstanceId $InstanceId | Out-Null
    throw "Expected rejection: $Name"
  }
  catch {
    if ($_.Exception.Message -notlike "*$ExpectedMessage*") {
      throw
    }
  }

  Write-Output "PASS: $Name"
}

$context = & $scriptPath `
  -ItemUrl $itemUrl `
  -ActionItemInstanceId $actionItemInstanceId | ConvertFrom-Json

if ($context.kpiId -ne $kpiId) {
  throw "The parser returned the wrong KPI ID."
}
if ($context.actionItemInstanceId -ne $actionItemInstanceId) {
  throw "The parser returned the wrong action-item instance ID."
}

$parsedSessionId = [guid]::Empty
if (
  -not [guid]::TryParse($context.sessionId, [ref] $parsedSessionId) -or
  $parsedSessionId.ToString("D")[14] -ne "4"
) {
  throw "The generated Agency session ID is not UUID v4."
}

$promptOptionIndex = [array]::IndexOf($context.arguments, "--prompt")
if ($promptOptionIndex -lt 0 -or $promptOptionIndex + 1 -ge $context.arguments.Count) {
  throw "The generated Agency arguments do not contain a prompt."
}
$prompt = [string] $context.arguments[$promptOptionIndex + 1]
foreach ($requiredText in @(
  "You are being launched from the Service 360 (S360) portal",
  "Use the 'Production' environment",
  $kpiId,
  $actionItemInstanceId,
  $context.originRemote
)) {
  if (-not $prompt.Contains($requiredText)) {
    throw "The generated resolver prompt is missing required context: $requiredText"
  }
}

foreach ($requiredArgument in @("--autopilot", "--assisted-approval", "--no-ask-user")) {
  if ($requiredArgument -notin $context.arguments) {
    throw "The generated Agency arguments are missing $requiredArgument."
  }
  if (
    [array]::IndexOf($context.arguments, "s360-remediation: s360-breeze") -lt 0 -or
    [array]::IndexOf($context.arguments, "--disable-mcp-server") -lt 0
  ) {
    throw "The generated Agency arguments do not isolate the S360 MCP server."
  }
}
Write-Output "PASS: exact S360 action item produces official resolver context"

$compoundInstanceId = "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0|windows|public"
$compoundContext = & $scriptPath `
  -ItemUrl $itemUrl `
  -ActionItemInstanceId $compoundInstanceId | ConvertFrom-Json
if ($compoundContext.actionItemInstanceId -ne $compoundInstanceId) {
  throw "The parser rejected or changed a valid compound action-item instance ID."
}
Write-Output "PASS: compound action-item instance ID accepted"

Assert-Rejected `
  -Name "unsupported host rejected" `
  -Url "https://example.com/blades/security?blade=KPI:$kpiId" `
  -InstanceId $actionItemInstanceId `
  -ExpectedMessage "Expected an S360 security item URL"

Assert-Rejected `
  -Name "missing KPI rejected" `
  -Url "https://vnext.s360.msftcloudes.com/blades/security?blade=Summary" `
  -InstanceId $actionItemInstanceId `
  -ExpectedMessage "valid KPI ID"

Assert-Rejected `
  -Name "unsafe instance ID rejected" `
  -Url $itemUrl `
  -InstanceId "value; Remove-Item ." `
  -ExpectedMessage "must be a single identifier"
