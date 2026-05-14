# Bot errors

The legacy detailed reference for bot-specific errors lives at [`docs/_v3-reference/bot-help.md`](../_v3-reference/bot-help.md). It is the source `helpLink` URLs in `aka.ms/teamsfx-bot-help` resolve to. This page indexes its contents and adds newer error names.

## Index of legacy reference

| Error name (`error.name`) | Where to look |
|---------------------------|--------------|
| `BT.MissingSubscriptionRegistrationError` | [bot-help.md → MissingSubscriptionRegistrationError](../_v3-reference/bot-help.md#btmissingsubscriptionregistrationerror) |
| `BT.BotRegistrationNotFoundError` | [bot-help.md → BotRegistrationNotFoundError](../_v3-reference/bot-help.md#btbotregistrationnotfounderror) |
| Reuse existing AAD | [bot-help.md → How to reuse existing AAD in Toolkit v2](../_v3-reference/bot-help.md#how-to-reuse-existing-aad-in-toolkit-v2) |

## Newer error names (v4 drivers)

| Error name | Driver | Mitigation |
|-----------|--------|-----------|
| `InvalidDriverInput` | any | Check the YAML key path in the error message; common cause: typo in `with:` block |
| `BotAadAppCreateFailed` | `botAadApp/create` | Re-run after `atk account login m365`; verify Graph permissions |
| `BotFrameworkRegisterFailed` | `botFramework/create` | Confirm the bot Entra ID app exists; verify ARM deployment succeeded |

## Common gotchas

- **Bot endpoint mismatch.** After re-provisioning, ensure `BOT_ENDPOINT` env value matches the actual App Service hostname. The `arm/deploy` driver outputs propagate; manual edits to the env file can fall out of sync.
- **Missing channel.** A bot needs both an Entra ID app *and* a Bot Channel registration. If you delete the channel manually, re-running provision will recreate it.
- **Two bots on one app.** A single Entra ID app can serve only one Bot Framework registration. Don't reuse a bot AAD app across projects.
