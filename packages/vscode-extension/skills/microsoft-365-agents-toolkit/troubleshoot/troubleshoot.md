# Troubleshooting

Consolidated troubleshooting for ATK projects — provisioning, runtime, Playground, and Teams issues.

## Common Provisioning Issues

| Symptom | Cause | Fix |
|---------|-------|-----|
| YAML schema validation error during `atk provision` | Wrong field names in `m365agents.yml` or `m365agents.local.yml` | Check [field reference](../references/manifest-and-yaml.md). Common: `outputJsonPath` → `outputFolder`, missing `description: ""` in `botFramework/create` |
| `teamsApp/validateManifest` fails with network error | Schema URL (`https://developer.microsoft.com/...`) unreachable | Remove `teamsApp/validateManifest` from local YAML, or retry with network access |
| `AADSTS7000229: missing service principal` | `aadApp/create` missing `generateServicePrincipal: true` | Add `generateServicePrincipal: true` to `aadApp/create` in YAML, re-provision — see [Missing Service Principal](#missing-service-principal-aadsts7000229) |
| 401 from Bot Connector (bot receives messages but can't reply) | `TENANT_ID` missing from `.localConfigs` → SDK uses wrong token authority | Copy `TENANT_ID` from `env/.env.local` to `.localConfigs` — see [Missing TENANT_ID](#missing-tenant_id-wrong-token-authority--401) |
| Bot still gets 401 after fixing auth issues | Devtunnel URL blacklisted by Bot Framework due to repeated prior failures | Create a fresh devtunnel (`devtunnel delete` + `devtunnel create`), update `BOT_ENDPOINT`, re-provision — see [Blacklisted Devtunnel URL](#blacklisted-devtunnel-url) |
| `Authorization: Bearer null` (401) at runtime | `clientId`/`clientSecret` not passed to Teams SDK `App` constructor | Pass credentials explicitly: `new App({ adapter: { credentials: { clientId, clientSecret, tenantId } } })` |
| 401 after changing to single-tenant (`AzureADMyOrg`) | Tenant mismatch — SDK doesn't accept `api://botid-{appId}` audience | Add custom JWT middleware accepting all audience formats, or stay with `AzureADMultipleOrgs` |
| Stale bot after re-provisioning | Old AAD app still referenced by Bot Framework registration | Delete `env/.env.local` and `env/.env.local.user`, re-run `atk provision --env local -i false` + `atk deploy --env local -i false` |
| Bot works in Playground but not in Teams | Missing dev tunnel or wrong `BOT_ENDPOINT` | Start `devtunnel host -p 3978 --allow-anonymous`, set `BOT_ENDPOINT` in `env/.env.local` before provisioning |
| Manifest v1.25 validation fails with `"team"` scope | `supportsChannelFeatures` required at runtime but rejected by v1.25 schema | Use `"personal"` scope only in v1.25, or use devPreview schema that defines the property |

## YAML Schema Errors

Common field name mistakes in `m365agents.local.yml`:
- `outputJsonPath` does not exist — use `outputFolder` in `teamsApp/zipAppPackage`
- `AAD_APP_OBJECT_ID` — use `BOT_OBJECT_ID` in local YAML's `aadApp/create` writeToEnvironmentFile
- Missing `description: ""` in `botFramework/create` — this field is required

See [../references/manifest-and-yaml.md](../references/manifest-and-yaml.md) for the full field reference.

## Known ATK Pitfalls

| Pitfall | Symptom | Fix |
|---------|---------|-----|
| `aadApp/create` missing `generateServicePrincipal: true` | `AADSTS7000229: missing service principal in tenant` when bot calls Bot Connector | Add `generateServicePrincipal: true` to `aadApp/create` in YAML, then re-provision |
| `TENANT_ID` not written to `.localConfigs` | SDK defaults to `botframework.com` tenant → 401 from Bot Connector (wrong issuer/tid in token) | Copy `TENANT_ID` from `env/.env.local` (where `aadApp/create` writes it) into `.localConfigs` |
| Devtunnel URL blacklisted after repeated 401s | Bot still gets 401 even after fixing auth — Bot Framework cached the tunnel URL as failing | Delete old tunnel, create a fresh one, update `BOT_ENDPOINT`, re-provision |

## Authorization / 401 Issues

### Missing Service Principal (AADSTS7000229)

The `aadApp/create` action in `m365agents.local.yml` must include `generateServicePrincipal: true` to create the service principal (enterprise application) alongside the app registration. Without it, the client credentials grant fails:

```
AADSTS7000229: The client application <BOT_ID> is missing service principal in the tenant <TENANT_ID>
```

**Fix — add `generateServicePrincipal: true` to your YAML:**
```yaml
  - uses: aadApp/create
    with:
      name: ${{CONFIG__MANIFEST__NAME}}-aad
      generateClientSecret: true
      generateServicePrincipal: true   # ← REQUIRED — without this, no SP is created
      signInAudience: AzureADMultipleOrgs
    writeToEnvironmentFile:
      clientId: BOT_ID
      clientSecret: SECRET_BOT_PASSWORD
      objectId: BOT_OBJECT_ID
```

Then re-provision:
```bash
atk provision --env local -i false
```

> **Manual fallback** (if you can't re-provision): `az ad sp create --id <BOT_ID>`

### Blacklisted Devtunnel URL

After repeated 401 failures (e.g., from a missing service principal), Bot Framework may blacklist the devtunnel URL. Even after fixing the auth issue, the bot continues to get 401.

**Fix — create a fresh devtunnel:**
```bash
devtunnel delete <old-tunnel-id>
devtunnel create --allow-anonymous
devtunnel port create -p 3978
devtunnel host
```

Update `BOT_ENDPOINT` in `env/.env.local` with the new tunnel URL, then re-provision:
```bash
atk provision --env local -i false
atk deploy --env local -i false
```

### Missing TENANT_ID (wrong token authority → 401)

When `TENANT_ID` is not set in `.localConfigs` or environment, the Teams SDK (both Python and Node) may default to acquiring tokens from the shared `botframework.com` tenant (`d6d49420-f39b-4df7-a1dc-d59a935871db`) instead of your home tenant. The resulting token has:
- Wrong `iss` (issuer) and `tid` (tenant) claims
- No `roles` assigned

Bot Connector rejects this token with **401 Unauthorized**.

**Diagnose:**
```bash
# Check if TENANT_ID is set
grep TENANT_ID .localConfigs
# Or in env file
grep TENANT_ID env/.env.local
```

**Fix:**
```bash
# Copy TENANT_ID from env file (aadApp/create writes it there, not to .localConfigs)
grep TENANT_ID env/.env.local
# Add the value to .localConfigs
echo TENANT_ID=<tenant-id-from-env-file> >> .localConfigs
```

This ensures the SDK uses `https://login.microsoftonline.com/<your-tenant-id>` instead of `https://login.microsoftonline.com/botframework.com`.

> **Python SDK note:** `TokenManager._resolve_tenant_id()` falls back to `botframework.com` when `TENANT_ID` is unset. Always set it explicitly.

### `Authorization: Bearer null`

The Teams SDK v2 `App` constructor requires explicit credentials. If `clientId`/`clientSecret` are not passed, the auth header will be `Bearer null`:
```typescript
const app = new App({
  adapter: {
    credentials: {
      clientId: process.env.CLIENT_ID,
      clientSecret: process.env.CLIENT_SECRET,
      tenantId: process.env.TENANT_ID, // required for single-tenant
    },
  },
});
```
Ensure `.localConfigs` has `CLIENT_ID` and `CLIENT_SECRET`. Run `atk deploy --env local -i false` to regenerate.

### 401 with single-tenant bots (`AzureADMyOrg`)

If `aadApp/create` uses `signInAudience: AzureADMyOrg`, Bot Framework tokens have audience `api://botid-{appId}`. The Teams SDK v2 only validates `{appId}` and `api://{appId}` by default, causing 401 errors. Solutions:
1. **Stay with `AzureADMultipleOrgs`** (recommended for most scenarios)
2. **Create custom auth middleware** with `skipAuth: true` on the `HttpPlugin`, then manually validate JWT tokens accepting all three audience formats: `{appId}`, `api://{appId}`, `api://botid-{appId}`

## Stale Bot Framework Registration

If you delete and re-create Azure AD apps, the Bot Framework registration may still reference the old app ID. Fix:
1. Delete `env/.env.local` and `env/.env.local.user`
2. Re-run `atk provision --env local -i false`
3. Re-run `atk deploy --env local -i false`
4. Re-sideload the Teams app

## Playground Issues

### Playground won't start

Check if port is in use:
```powershell
# Windows
netstat -ano | findstr :56150

# The playground will automatically find an available port if 56150 is in use
```

### Bot not responding in Playground

1. Verify bot is running on specified endpoint
2. Check bot logs for errors
3. Ensure your bot endpoint is accessible:
   ```bash
   curl http://localhost:3978/api/messages
   ```

## Teams Issues

### Teams shows "app not available"

This usually means BOT_ENDPOINT requires HTTPS. Use Agents Playground instead, or ensure dev tunnel is running and BOT_ENDPOINT is properly configured.

### App not loading

Verify `M365_APP_ID` (for declarative agents) or `TEAMS_APP_ID` (for bots/tabs) exists in `env/.env.local`.

## Runtime Issues

### Port already in use

```powershell
Get-NetTCPConnection -LocalPort 3978 | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

### Missing environment variables at runtime

Check `.localConfigs` exists and has required values. Run `atk deploy --env local -i false` to regenerate.

## Diagnostics Commands

```bash
atk doctor              # Check ATK installation and dependencies
atk validate --env <env> # Validate project configuration
atk auth list           # Check logged-in accounts
agentsplayground --help # Playground CLI help
atk provision --help    # Provision help
atk deploy --help       # Deploy help
```
