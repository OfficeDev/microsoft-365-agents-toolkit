# Scenario - Create Declarative Agent with Office Add-in Action (`declarative-agent-meta-os-new-project`)

- **Status:** Accepted (v4 migration of the v3 MetaOS new-project create template) - ready for scenario-tier (T3) tests
- **Domain:** [`01-scaffolding`](../../domains/01-scaffolding.md)
- **Scenario ID:** `SCN-DA-CREATE-METAOS-NEW-PROJECT`
- **Template id:** `declarative-agent-meta-os-new-project` (create)

This is the vertical contract for the native v4 Declarative Agent with Office Add-in Action new-project package. The v3 path renders the static `declarative-agent-meta-os-new-project` template, then runs `MetaOSHelper.unifyProjectID`: set the Teams manifest id and `env/.env.dev` `TEAMS_APP_ID` to the same generated UUID. The v4 package preserves that behavior as a post-render `metaos/unify-project-id` pipeline step.

## Acceptance Criteria

| ID | Tier | Given | When | Then |
|----|------|-------|------|------|
| SCN-CREATE-METAOS-01 | L1 | empty target | scaffold completes | the render phase writes the same new-project file set as v3 (`.tpl` stripped), including Office Add-in source, DA appPackage, infra, evals, env, and yaml files |
| SCN-CREATE-METAOS-02 | L1 | rendered `appPackage/manifest.json` and `env/.env.dev` | `metaos/unify-project-id` runs | manifest `id` and env `TEAMS_APP_ID` are both set to the same generated UUID, replacing the rendered placeholder / empty env value |
| SCN-CREATE-METAOS-03 | L1 | rendered `appPackage/declarativeAgent.json` and `appPackage/alchemy-plugin.json` | render | DA and action manifests match the v3 template shape: DA action `alchemyPlugin` references `alchemy-plugin.json`; the action manifest contains the Office runtime function definitions |
| SCN-CREATE-METAOS-04 | L1 | empty target | scaffold | pipeline steps run in order: `require-empty-target`, then `metaos/unify-project-id` |
| SCN-CREATE-METAOS-05 | L1 | non-empty target | scaffold | `require-empty-target` fails first with **`UserError`** and writes nothing |

## Composed operations

- [`walk-create-selector`](../../operations/scaffolding/walk-create-selector.md) - routes DA add-action MetaOS to this v4 package when `TEAMSFX_DA_METAOS` is enabled.
- [`run-scaffold-pipeline`](../../operations/scaffolding/run-scaffold-pipeline.md) - runs `require-empty-target`, renders the template, then executes [`metaos/unify-project-id`](../../operations/scaffolding/metaos-unify-project-id.md).

## Boundary

This scenario covers only the new-project MetaOS create path. It does **not** cover `declarative-agent-meta-os-upgrade-project`, which copies and mutates an existing Office Add-in project and needs a separate modify/import operation.