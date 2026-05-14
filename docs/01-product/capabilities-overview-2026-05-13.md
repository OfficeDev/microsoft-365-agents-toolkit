# Microsoft 365 Agents Toolkit — Capability Overview

**Date:** 2026-05-13  
**Purpose:** Understand the current toolkit surface area as input for v4 design. Organized by capability domain to map what exists, where boundaries are clear, and where v4 needs to improve reliability and maintainability.

**v4 Core Goal:** Make AI coding more reliable by establishing clear contracts between components, enabling Design-First TDD: Design → Write Tests → Implement → Run Tests (loop).

---

## Entry Points

The toolkit exposes capabilities through three entry points:

| Entry Point | Description |
|---|---|
| **CLI** (`atk`) | Command-line interface for all lifecycle operations |
| **VS Code Extension** | Full IDE integration with tree views, debugging, deployment |
| **Core Engine API** | `core-next` TypeScript modules consumed by CLI and extension |

---

## Capability Domains

### Domain 1: Scaffolding

**What users can do:** Create a new project from templates or existing samples, selecting app type, language, and authentication.

**Entry points:**
| | Command |
|---|---|
| CLI | `atk new --capability <type> --app-name <name> [options]` |
| CLI (sample) | `atk new sample <sample-name>` |
| VS Code | `fx-extension.create`, `fx-extension.openSamples` |

**Core module:** `core-next/src/templates` + `core-next/src/questions`

**Template catalog:** 40+ templates across 6 project types:
- Declarative Agent (11 templates, including MCP, TypeSpec, Graph Connector)
- Custom Engine Agent (3 templates)
- Teams Agent (7 templates, including RAG variants)
- Tab (6 templates, including SPFx, SSO, Dashboard)
- Bot (8 templates, including notification, workflow, command-response)
- Message Extension (8 templates)
- Office Add-in (4 templates)

**Key note:** `atk new` uses a flags-based interface (`--capability`, `--with-plugin`, `--llm-service`, etc.), not subcommands. With `--interactive true` (default), it guides users through prompts. In non-interactive mode (CI / AI coding), the flag surface is large and complex.

**Asymmetry:** Sample browsing/creation is available in VS Code (`openSamples`) and via `atk new sample`. Template creation is available via both CLI and VS Code.

**Current state:**
| | Status |
|---|---|
| Spec clarity | ⚠️ Template names and categories defined; question/selection flow has no formal spec |
| Test derivability | ⚠️ "Generated project structure" tests derivable from template manifest, but coverage is sparse |
| Boundary clarity | ✅ v4 has separated the template system from core logic |

---

### Domain 2: Lifecycle

**What users can do:** Move a project from local code to cloud-deployed — provisioning M365 resources, deploying code, publishing to the Teams app catalog, and running locally or on Agents Playground.

**Toolkit's core differentiator:** `provision` hides the complexity of creating M365-related resources that are otherwise scattered across multiple portals and require manual, error-prone steps:

| M365 Resource | What it does |
|---|---|
| Teams App registration | Creates/updates app in Developer Portal |
| Entra App registration | Creates Azure AD app, configures permissions |
| OAuth registration | Configures Teams OAuth connection |
| API Key registration | Registers API Key provider |
| Upload App Package | Packages and uploads manifest zip to Teams/M365 |

> **Note:** OAuth registration and API Key registration details need verification against actual driver implementations.

Azure resources (Bot Service, Storage, Functions, etc.) are optional additions — not the core value.

**Two paths:**

```
Path A (requires Provision)            Path B (no Provision needed)
─────────────────────────────          ──────────────────────────────
Provision (register M365 resources)    Run directly
      ↓
Local Debug          ← ★ KPI           Agents Playground   ← ★ KPI
Remote Preview                         DevTools
```

**KPI priority:**
1. **Run on Agents Playground** — lightest path, no cloud resources needed, fastest validation loop
2. **Local Debug** — requires provision, but core development experience

**Entry points:**
| | Command |
|---|---|
| CLI | `atk provision`, `atk deploy`, `atk publish`, `atk preview` |
| VS Code | `fx-extension.provision`, `fx-extension.deploy`, `fx-extension.publish`, `fx-extension.localdebug`, `fx-extension.preview` |

**Core module:** `core-next/src/lifecycle` + `core-next/src/drivers`

**Current state:**
| | Status |
|---|---|
| Spec clarity | ⚠️ Two paths not formally distinguished in documentation |
| Test derivability | ❌ Path A depends on real M365 tenant; Path B relatively testable but under-covered |
| Boundary clarity | ⚠️ State between steps flows through `.env.*` files implicitly |

---

### Domain 3: Identity

**What users can do:** Log in and out of M365 and Azure accounts, switch tenants, view current authentication status.

**Entry points:**
| | Command |
|---|---|
| CLI | `atk auth login azure`, `atk auth login m365`, `atk auth logout`, `atk auth show` |
| VS Code | `fx-extension.cmpAccounts`, `fx-extension.signOut`, `fx-extension.m365SwitchTenant`, `fx-extension.azureSwitchTenant` |

**Core module:** `core-next/src/clients`

**Two account systems:**

| Account | Required when |
|---|---|
| **M365 account** | Depends on provision actions — e.g., registering Teams App, Entra App |
| **Azure account** | Depends on provision actions — e.g., creating Azure resources |

**Key principle:** Account requirements are determined entirely by which provision actions are defined in the project, not at the domain level. Some actions require M365, some require Azure, some require both.

**Current state:**
| | Status |
|---|---|
| Spec clarity | ⚠️ Each client has its own auth flow; no unified token lifecycle spec |
| Test derivability | ❌ Requires real accounts; mocking is difficult |
| Boundary clarity | ⚠️ How tokens are passed to downstream drivers is implicit |

---

### Domain 4: Manifest

**What users can do:** Validate, package, upload, and sync Teams App manifest; manage Entra App manifest.

**Entry points:**
| | Command |
|---|---|
| CLI | `atk teamsapp validate`, `atk teamsapp package`, `atk teamsapp publish`, `atk teamsapp update` |
| VS Code | `fx-extension.validateManifest`, `fx-extension.updatePreviewFile`, `fx-extension.updateAadAppManifest`, `fx-extension.syncManifest`, `fx-extension.publishInDeveloperPortal` |

**Core module:** `core-next/src/manifest` + `packages/manifest`

**Two manifest types:**

| Type | Description |
|---|---|
| **Teams App Manifest** | Defines app capabilities, permissions, entry points; packaged as zip for upload |
| **Entra App Manifest** | Defines Azure AD app permissions and configuration |

**Current state:**
| | Status |
|---|---|
| Spec clarity | ⚠️ Manifest format follows official Microsoft schema; but Toolkit's template variable substitution logic (`${{VAR}}`) has no formal spec |
| Test derivability | ✅ validate/package logic is purely local — best candidate for Design-First TDD |
| Boundary clarity | ⚠️ Manifest variables depend on `.env.*` injection, coupling this domain to Domain 6 |

**Best opportunity:** This domain has the clearest local logic and the least external dependencies — ideal starting point for Design-First TDD.

---

### Domain 5: Extensibility

**What users can do:** Add new capabilities to an existing project — actions (API/MCP), knowledge sources, agent capabilities, or regenerate action resources.

**Entry points:**
| | Command |
|---|---|
| CLI | `atk add action`, `atk add capability`, `atk add auth-config` |
| VS Code | `fx-extension.addPlugin`, `fx-extension.addKnowledge`, `fx-extension.addWebpart`, `fx-extension.regeneratePlugin`, `fx-extension.metaOSExtendToDA` |

**Core module:** `core-next/src/declarativeAgent`

**Three extension types:**

| Type | Description |
|---|---|
| **Action** | Add API plugin (via OpenAPI spec) or MCP server to Declarative Agent |
| **Knowledge** | Add knowledge source (Graph Connector, SharePoint, etc.) |
| **Capability** | Extend agent capabilities (Code Interpreter, Image Generator, etc.) |

**Current state:**
| | Status |
|---|---|
| Spec clarity | ⚠️ Difference between OpenAPI vs MCP action addition flows not formally spec'd |
| Test derivability | ⚠️ Depends on spec parser for OpenAPI; testable but under-covered |
| Boundary clarity | ❌ A single `add action` touches manifest, scaffolding, and provision simultaneously — side effects are scattered |

**Key risk for AI coding:** One `add action` operation has cross-domain side effects (manifest + scaffolding + provision). AI cannot reliably infer the complete impact of a single operation. This is the hardest domain for reliable AI-generated code.

---

### Domain 6: Environment Management

**What users can do:** Create and manage multiple deployment environments (dev, staging, prod), each with independent resource state.

**Entry points:**
| | Command |
|---|---|
| CLI | `atk env add`, `atk env list`, `atk env reset` |
| VS Code | `fx-extension.addEnvironment` |

**Core module:** `core-next/src/environment`

**How it works:** An environment is a set of `.env.<name>` files storing the output state of all provision actions for that environment.

```
.env.dev        → dev environment resource state
.env.staging    → staging environment resource state
.env.local      → local debug only
```

**The driver schema system (already exists):**

Each driver declares:
- **Inputs:** Zod schema (`inputSchema`) — validated at compile/runtime
- **Outputs:** `writeToEnvironmentFile` in YAML — maps driver output keys to env var names
- **YAML schema:** Defines required `writeToEnvironmentFile` keys per driver

Example (aadApp/create):
```yaml
- uses: aadApp/create
  with:
    name: MyApp
    generateClientSecret: true
  writeToEnvironmentFile:
    clientId: AAD_APP_CLIENT_ID      # driver key → env var name
    objectId: AAD_APP_OBJECT_ID
    clientSecret: SECRET_AAD_APP_CLIENT_SECRET
```

**The remaining problem:** Schema exists at the YAML layer (runtime), but not at the TypeScript layer (compile time). `DriverOutput` is typed as `Record<string, string>` — output keys are invisible to the TypeScript compiler and to AI coding agents reading TypeScript.

**Proposed improvement (no user behavior change):** Lift output keys to TypeScript generics:

```typescript
// Before
interface DriverOutput {
  outputs: Record<string, string>;
}

// After
interface DriverOutput<TOutputs extends string = string> {
  outputs: Record<TOutputs, string>;
}

type AadAppCreateOutputs = "clientId" | "objectId" | "tenantId" | "authority" | "authorityHost" | "clientSecret";

createDriver<AadAppCreateConfig, AadAppCreateOutputs>({ ... })
```

**Effect:** AI can statically infer env state from TypeScript code alone, without reading YAML schema. Compile-time enforcement of output key correctness. Zero user-visible change.

**Current state:**
| | Status |
|---|---|
| Spec clarity | ✅ YAML schema defines per-driver output contracts |
| Test derivability | ⚠️ env add/list testable; reset involves cloud state |
| Boundary clarity | ⚠️ TypeScript output types are untyped (`Record<string, string>`) |

---

### Domain 7: Collaboration

**What users can do:** Share M365 resource access with team members, manage collaborator lists.

**Entry points:**
| | Command |
|---|---|
| CLI | `atk permission grant`, `atk permission status` |
| VS Code | `fx-extension.grantPermission`, `fx-extension.listCollaborator`, `fx-extension.manageCollaborator` |

**Core module:** `core-next/src/clients` (Graph API, Teams Dev Portal clients)

**What it does:** Grants Entra App owner permissions and Teams App management rights to specified M365 user accounts.

**Current state:**
| | Status |
|---|---|
| Spec clarity | ⚠️ Permission scope (what exactly is granted) not formally defined |
| Test derivability | ❌ Requires real M365 tenant and user accounts |
| Boundary clarity | ✅ Relatively isolated; operates via Graph API with clear side-effect boundary |

**Dependencies:** Relies on Domain 3 (Identity) for tokens, Domain 6 (Environment) for stored App IDs.

---

## Global Cross-Domain Analysis

### Status Summary

| Domain | Spec Clarity | Test Derivability | Boundary Clarity | v4 Coverage |
|---|---|---|---|---|
| 1. Scaffolding | ⚠️ | ⚠️ | ✅ | ✅ Rebuilt |
| 2. Lifecycle | ⚠️ | ❌ | ⚠️ | ⚠️ Partial |
| 3. Identity | ⚠️ | ❌ | ⚠️ | ⚠️ Partial |
| 4. Manifest | ⚠️ | ✅ | ⚠️ | ⚠️ Partial |
| 5. Extensibility | ⚠️ | ⚠️ | ❌ | ⚠️ Partial |
| 6. Environment | ✅ (YAML) | ⚠️ | ⚠️ | ❌ Unresolved |
| 7. Collaboration | ⚠️ | ❌ | ✅ | ❌ Not rebuilt |

### Systemic Risk: Implicit `.env.*` State

All domains share implicit state through `.env.*` files. While the driver schema system provides YAML-layer contracts, the TypeScript layer has no visibility into what keys exist after a given set of actions. This is the single issue that most broadly affects AI coding reliability across all domains.

### Cross-Domain Side Effects (Domain 5)

`add action` is the highest-risk operation: it simultaneously modifies manifest (Domain 4), generates scaffold files (Domain 1), and may trigger provision registration (Domain 2). No single domain owns this operation, making it the hardest to test and the most fragile under AI-driven changes.

### Design-First TDD Starting Point

Domain 4 (Manifest) is the best entry point for establishing the Design-First TDD workflow:
- Purely local logic (no cloud dependencies)
- Clear official schema to derive tests from
- High user-visible impact (validation errors surface immediately)

---

## v4 Priority Recommendations

```
Immediate (unblocks other domains)
  └── Domain 6: Lift driver output keys to TypeScript generics
                → Zero user behavior change, maximum AI reliability gain

High priority (KPI-direct)
  ├── Domain 2: Formally spec both paths (Provision vs. Agents Playground)
  └── Domain 4: Derive complete test suite from manifest spec (Design-First TDD pilot)

Medium priority
  ├── Domain 5: Decompose add-action cross-domain side effects into explicit steps
  └── Domain 3: Declare account requirements per action (not per domain)

Lower priority
  └── Domain 7: Isolated domain, can be addressed later
```

---

## Items Needing Verification

- [ ] OAuth registration and API Key registration: exact behavior and driver implementation details
- [ ] Agents Playground path: confirm exact entry point and whether any M365 resources are needed
