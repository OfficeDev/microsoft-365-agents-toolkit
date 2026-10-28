param(
  [Parameter(Mandatory)]
  [ValidateSet("static-oauth", "dcr", "entra-sso")]
  [string]$Mode,
  [string]$SsoEnvironmentFile = "",
  [string]$StaticCredentialFile = "",
  [ValidateSet("auto", "absent", "invalid")]
  [string]$SsoVariant = "auto",
  [string]$AccountProfile = "atk08",
  [string]$Image = "vscuse-atk-feat-16840-dcr-update:local",
  [switch]$DcrUpdate,
  [switch]$Author
)

$ErrorActionPreference = "Stop"
$Mode = $Mode.ToLowerInvariant()
$SsoVariant = $SsoVariant.ToLowerInvariant()
$repo = (git -C $PSScriptRoot rev-parse --show-toplevel).Trim()
$previous = @{}
Get-ChildItem Env: | ForEach-Object { $previous[$_.Name] = $_.Value }
try {
  if ($Mode -ne "entra-sso" -and $SsoVariant -ne "auto") {
    throw "SSO registration variants are only supported by the SSO case."
  }
  if ($DcrUpdate -and $Mode -ne "dcr") {
    throw "The independent DCR update case requires Mode dcr."
  }
  . (Join-Path $repo "set-azure-env.ps1") -Env $AccountProfile *> $null
  if (-not $SsoEnvironmentFile -and $StaticCredentialFile) {
    throw "File-based static credentials require the matching consumer environment file."
  }

  $helper = Join-Path $PSScriptRoot "engine\prepare-mcp-auth-run.cjs"
  $profileArguments = ConvertTo-Json -InputObject @{
    mode = $Mode
    ssoFile = $SsoEnvironmentFile
    staticFile = $StaticCredentialFile
    variant = $SsoVariant
  } -Compress
  $profileArgumentsBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($profileArguments))
  $encodedProfile = & node -e 'const h=require(process.argv[1]);const a=JSON.parse(Buffer.from(process.argv[2],"base64"));const p=a.ssoFile?h.loadProfile(a.mode,a.ssoFile,a.staticFile,a.variant):h.loadEnvironmentProfile(a.mode,process.env,a.variant);process.stdout.write(JSON.stringify(p.environment));' $helper $profileArgumentsBase64
  if ($LASTEXITCODE -ne 0) { throw "MCP fixture profile validation failed." }
  try { $environment = $encodedProfile | ConvertFrom-Json -AsHashtable }
  catch { throw "MCP fixture profile could not be decoded safely." }
  Get-ChildItem Env: |
    Where-Object { $_.Name -match "^(MCP_AUTH_|DA_MCP_|ENTRA_SSO_|LOCAL_CLIENT_|STATIC_API_TOKENS$)" } |
    ForEach-Object { [Environment]::SetEnvironmentVariable($_.Name, $null, "Process") }
  foreach ($name in $environment.Keys) {
    [Environment]::SetEnvironmentVariable($name, [string]$environment[$name], "Process")
  }
  $env:M365_TENANT_ID = $env:MCP_AUTH_EXPECTED_TENANT
  if ($env:M365_ACCOUNT_NAME -notmatch "@atktesting08\.onmicrosoft\.com$") {
    throw "The selected M365 account is not in the requested atktesting08 profile."
  }
  foreach ($name in @("M365_ACCOUNT_PASSWORD", "AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_API_KEY", "AZURE_VISION_ENDPOINT", "AZURE_VISION_KEY")) {
    if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($name))) {
      throw "Missing required runner setting: $name"
    }
  }
  $env:VSCUSE_VSCODE_IMAGE = $Image
  $env:NPM_CONFIG_REGISTRY = "https://packagefeedproxy.microsoft.io/npm/"
  $env:PYTHONUNBUFFERED = "1"
  $runnerDirectory = Join-Path $env:USERPROFILE ".local\vscuse\venv\Scripts"
  $env:PATH = "C:\Program Files\Docker\Docker\resources\bin;$runnerDirectory;$env:PATH"
  $existingContainer = & docker ps -aq --filter "name=^/vscuse-container$"
  if ($LASTEXITCODE -ne 0) { throw "Docker is not ready for an isolated MCP test." }
  if ($existingContainer) {
    throw "An existing vscuse container must be closed by its owner before starting an independent authentication case."
  }
  $caseNames = @{
    "static-oauth" = "da-mcp-static-entra-copilot"
    "dcr" = "da-mcp-dcr-copilot"
    "entra-sso" = "da-mcp-entra-sso-copilot"
  }
  if ($DcrUpdate) { $caseNames["dcr"] = "da-mcp-dcr-update-copilot" }
  $plan = Join-Path $PSScriptRoot ("plans\da-mcp-server--" + $caseNames[$Mode] + ".json")
  $runName = "{0}-{1}-{2}" -f (Get-Date -Format "yyyyMMdd-HHmmss"), $caseNames[$Mode], $SsoVariant
  $reportDirectory = Join-Path $repo (".local\test-reports\" + $runName)
  New-Item -ItemType Directory -Force -Path $reportDirectory | Out-Null
  $config = Join-Path $reportDirectory "config.yaml"
  $names = ConvertTo-Json -InputObject @($environment.Keys) -Compress
  $ssoIdPresent = $Mode -eq "entra-sso" -and -not [string]::IsNullOrEmpty([string]$environment["MCP_AUTH_SSO_ID"])
  & node -e 'const fs=require("fs"),h=require(process.argv[1]);const text=h.prepareMcpAuthRunConfig(fs.readFileSync(process.argv[2],"utf8"),fs.readFileSync(process.argv[3],"utf8"),process.argv[4],JSON.parse(process.argv[5]),JSON.parse(process.argv[7]));fs.writeFileSync(process.argv[6],text);' $helper (Join-Path $PSScriptRoot "config.yaml") $plan $Mode $names $config ($ssoIdPresent.ToString().ToLowerInvariant())
  if ($LASTEXITCODE -ne 0) { throw "Mode-scoped vscuse configuration failed." }
  $secrets = @(Get-ChildItem Env: | Where-Object { $_.Name -match "SECRET|PASSWORD|(^|_)KEY$|TOKEN|MCP_AUTH_SSO_ID|MCP_AUTH_SSO_PROTECTED_ID_B64|MCP_AUTH_AAD_MANIFEST_B64" } | ForEach-Object { $_.Value } | Where-Object { $_.Length -ge 4 })
  $secretVariants = @($secrets | ForEach-Object {
    $_
    [uri]::EscapeDataString($_)
    [System.Net.WebUtility]::HtmlEncode($_)
    $encoded = ConvertTo-Json -InputObject $_ -Compress
    $encoded.Substring(1, $encoded.Length - 2)
  } | Sort-Object -Unique)
  $log = Join-Path $reportDirectory "run.log"
  Push-Location $PSScriptRoot
  try {
    $runner = if ($Author) { Join-Path $runnerDirectory "vscuse-ui.exe" } else { Join-Path $runnerDirectory "vscuse.exe" }
    $arguments = if ($Author) {
      # Generated MCP plans are flat; no shared group lookup is required.
      $authoringDirectory = Join-Path $reportDirectory "authoring"
      New-Item -ItemType Directory -Force -Path $authoringDirectory | Out-Null
      Copy-Item -LiteralPath $plan -Destination $authoringDirectory
      @("--config-file", $config, "--project-path", $authoringDirectory)
    } else {
      @("execute", "--config-file", $config, "--groups-dir", "groups", "--report-dir", $reportDirectory, $plan)
    }
    & $runner @arguments 2>&1 | ForEach-Object {
      $line = $_.ToString()
      foreach ($secret in $secretVariants) { $line = $line.Replace($secret, "[REDACTED]") }
      $line -replace "\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b", "[REDACTED_TOKEN]" -replace "([?&](?:code|access_token|refresh_token|id_token|client_secret)=)[^&\s`"'<>]+", '$1[REDACTED]'
    } | Tee-Object -FilePath $log
    $exitCode = $LASTEXITCODE
    foreach ($file in Get-ChildItem -LiteralPath $reportDirectory -Recurse -File | Where-Object { $_.Extension -in @(".html", ".json", ".log") }) {
      $text = [IO.File]::ReadAllText($file.FullName)
      foreach ($secret in $secretVariants) { $text = $text.Replace($secret, "[REDACTED]") }
      $text = $text -replace "\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b", "[REDACTED_TOKEN]"
      $text = $text -replace "([?&](?:code|access_token|refresh_token|id_token|client_secret)=)[^&\s`"'<>]+", '$1[REDACTED]'
      [IO.File]::WriteAllText($file.FullName, $text)
    }
    Write-Output "MCP auth run artifacts: $reportDirectory"
    if ($exitCode -ne 0) { throw "The MCP UI run did not pass. Inspect its redacted run artifacts." }
  } finally { Pop-Location }
} finally {
  Get-ChildItem Env: | Where-Object { -not $previous.ContainsKey($_.Name) } |
    ForEach-Object { [Environment]::SetEnvironmentVariable($_.Name, $null, "Process") }
  foreach ($name in $previous.Keys) {
    [Environment]::SetEnvironmentVariable($name, $previous[$name], "Process")
  }
}
