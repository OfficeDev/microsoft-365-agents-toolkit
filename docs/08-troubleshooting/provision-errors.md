# Provision errors

Errors surfaced during the `provision` lifecycle.

## `InvalidDriverInput`

**Trigger.** A YAML action's `with:` block fails Zod validation.

**Mitigation.** The error message includes the Zod issue path (e.g. `with.subscriptionId: required`). Open `m365agents.yml`, find the failing step, fix the typo or supply the missing field. Re-run.

## `MissingResourceGroup`

**Trigger.** `ensureResourceGroup` fails because `AZURE_RESOURCE_GROUP_NAME` is set to an empty string.

**Mitigation.** Either remove the empty `AZURE_RESOURCE_GROUP_NAME=` line from your env file (lets the toolkit prompt with the default), or fill in a real RG name.

## `SubscriptionNotFound`

**Trigger.** The signed-in Azure account has no subscriptions accessible, or the named subscription is invalid.

**Mitigation.**

1. `atk account login azure` to confirm the account.
2. `az account list` to see subscriptions.
3. Set `AZURE_SUBSCRIPTION_ID` in the env file to one of the IDs.

## `ArmDeploymentFailed`

**Trigger.** ARM/Bicep deployment failed. Real cause is an inner ARM error — quota, naming conflict, permission, etc.

**Mitigation.** Re-run with `--debug` to see the ARM deployment detail. Common causes:

- Resource name collision in the same RG → change `appName` or delete the conflicting resource.
- vCPU quota → request a quota increase or switch region.
- Insufficient role → grant `Contributor` on the RG to the signed-in user / SP.

## `TeamsAppCreateFailed` (TDP 4xx)

**Trigger.** Manifest validation rejected by the Teams Developer Portal.

**Mitigation.** The error includes TDP's response. Fix the manifest field it complains about. Common: invalid icon size, malformed `validDomains`, missing `bots[].botId`.

## `M365TokenAcquisitionFailed`

**Trigger.** MSAL silent refresh failed and interactive flow is unavailable (CI mode).

**Mitigation in CI:** ensure the service principal is configured and the M365 app has the required Graph + Teams App scopes. In interactive mode this should re-prompt automatically.

## `Forbidden` (Graph 403)

**Trigger.** Signed-in user lacks permission to perform the operation (e.g. AAD app creation, app catalog publish).

**Mitigation.** Sign in as a user with `Application.ReadWrite.All` (for AAD app reg) or `AppCatalog.ReadWrite.All` (for publish). Or have an admin grant the toolkit's MSAL client these scopes.

## Provision is idempotent

Re-running provision after a failure is safe. Drivers check for existing resources via tracked IDs (`existingTeamsAppId`, `BOT_ID`, etc.) and update rather than recreate. To force a fresh state, run `atk env reset --env <name>` first.
