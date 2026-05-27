# Identity & Login — Code Map

Navigation aid for refactor work on the identity / login substrate. Maps each
fact in [`identity-and-login.md`](identity-and-login.md) to its current
location in source.

> **This file is not part of the contract.** It is expected to churn as code
> moves. Constraints live in
> [`identity-and-login.md`](identity-and-login.md#2-constraints-derived-from-these-facts);
> updates here do not require an ADR.

| Fact (from `identity-and-login.md` §1) | File(s) |
|---|---|
| §1.1 First-party client ID, broker config (VS Code) | `packages/vscode-extension/src/commonlib/m365Login.ts` |
| §1.1 First-party client ID, broker config (CLI) | `packages/cli/src/commonlib/m365Login.ts` |
| §1.2 / §1.4 Code-flow / broker / redirect plumbing | `packages/vscode-extension/src/commonlib/codeFlowLogin.ts`, `packages/cli/src/commonlib/codeFlowLogin.ts` |
| §1.2 VS Code Azure session bridge | `packages/vscode-extension/src/commonlib/vscodeAzureSubscriptionProvider.ts`, `packages/vscode-extension/src/commonlib/azureLogin.ts` |
| §1.1 / §1.5 Sovereign environment + Entra authority (`TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT`, `SovereignCloudEnvironment` enum) | `packages/fx-core/src/common/accountUtils.ts` |
| §1.5 / §1.6 Endpoint matrix + scope helpers | `packages/fx-core/src/common/constants.ts` |
| §1.5 / §1.7 Internal-dogfood TDP selector (`APP_STUDIO_ENV=int`) | `packages/fx-core/src/common/constants.ts`, `packages/fx-core/src/types/env.d.ts` |
| §1.7 TDP region discovery | `packages/fx-core/src/client/teamsDevPortalClient.ts` |
| §1.3 Broker feature flag (`TEAMSFX_BROKER_AUTH`, `FeatureFlags.BrokerAuth`) | `packages/fx-core/src/common/featureFlags.ts` |
| §1.1 / §1.4 Reply-URL identifiers | `packages/vscode-extension/src/commonlib/common/constant.ts` |
