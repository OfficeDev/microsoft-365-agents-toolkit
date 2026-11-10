# SPFx scaffolding tooling — Code Map

SPFx scaffolding and its Yeoman/dependency-checker implementation have
been removed. See the
[retirement contract](../../03-specs/operations/product/retire-spfx-generation.md).

> **This file is not part of the contract.** It is expected to churn as code
> moves. Constraints live in
> [`spfx-tooling.md`](spfx-tooling.md#2-constraints-derived-from-these-facts);
> updates here do not require an ADR.

| Retained behavior | File(s) |
|---|---|
| Rejection-only legacy `addWebpart` API | `packages/fx-core/src/core/FxCore.ts` |
| Existing-project deployment | `packages/fx-core/src/component/driver/deploy/spfx/deployDriver.ts` |
| Existing-project app package copying | `packages/fx-core/src/component/driver/teamsApp/copyAppPackageToSPFx.ts` |

> §1.4 (`@microsoft/spfx-cli` / `@microsoft/spfx-template-api`) is an
> upstream-owned fact and has no corresponding toolkit source today; it
> no replacement binding is planned. The
> [migration proposal](../adr/ADR-0009-spfx-scaffolding-tooling-path.md)
> is withdrawn.
