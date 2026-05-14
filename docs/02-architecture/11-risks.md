# 11 — Risks and technical debt

## Active risks

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|-----------|
| Two engines (v3 / v4) drift | Medium | High | Feature parity tests in CI · shared driver IDs · `features.json` integration tests |
| Manifest schema future versions break wrappers | Medium | High | Single typed wrapper layer (`@microsoft/app-manifest`) · schema-version pinning in templates |
| MSAL native plugins (keytar, msal-node-extensions) installation failures on user machines | Medium | Medium | Optional deps; graceful fallback to unencrypted cache + AES-256-GCM at rest |
| Long-tail of one-off generators in `fx-core` blocks v4 cutover | High | Medium | Each generator gets a v4 `TemplateDescriptor` equivalent; gap tracked in `features.json` |
| `swagger2openapi 7.0.8` exact pin (used by spec-parser) | Low | Medium | Pinned because newer versions break `RemoteMCPServer`; revisit on next OpenAPI bump |
| `typeSpec/compile` driver missing for `declarative-agent-typespec` template | Low | Low | Tracked; will be added when TypeSpec scaffold is widely used |

## Technical debt

| Debt | Why it accumulated | Plan |
|------|-------------------|------|
| `cli` v3 webpack config (4 GB heap, 100+ LOC) | Predates esbuild adoption | Retire when v4 reaches parity; users migrate via flag |
| `mcp-server` and `vscode-ui` still on webpack | Lower priority than CLI / extension | Migrate when next maintained |
| `secretmasker` v3 SVM + BloomFilter complexity | Predates the keyword-based masker | v4 keyword-only is enough; v3 stays for compatibility |
| `TOOLS` global singleton in `fx-core` | Pre-DI design | Replaced by `AtkContext` in v4 — not removed in v3 |
| 175 `no-explicit-any` warnings in `core-next`, 32 in `cli-next` | Tracked, not blocking; CI allows warnings | Removed opportunistically in PRs that touch them |

## Removed risks

- **VS Code extension build size.** Was a webpack risk; resolved by esbuild migration.
- **NodeJS 16 EoL.** Mitigated by targeting Node 18.

## Re-assessment cadence

This page is reviewed at the start of each release cycle. ADR `0001` (v4 feature flag) tracks the cutover — when it moves from "Accepted" to "Superseded", roughly half of these risks dissolve.
