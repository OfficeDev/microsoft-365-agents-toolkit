# 1 — Introduction and goals

## Requirements overview

The toolkit must enable a developer to go from **install** → **published Microsoft 365 app** without leaving their editor or shell. Detailed product requirements are in [01-product/prd-overview.md](../01-product/prd-overview.md).

## Top quality goals

| # | Goal | Why it matters | Where measured |
|---|------|----------------|----------------|
| 1 | **Reproducibility** of provision/deploy/publish across local + CI | Pipelines depend on this | Lifecycle success-rate telemetry |
| 2 | **Backward compatibility** of v3 projects | Tens of thousands of existing projects | Integration tests against v3 manifest schemas |
| 3 | **Schema fidelity** to Microsoft 365 manifest standards | Manifest is the contract with the platform | `@microsoft/app-manifest` typed wrapper |
| 4 | **Security** of credentials, code, and uploads | Token cache + Zip-Slip + magic bytes + PII masking | CodeQL + unit tests |
| 5 | **Localisation** in 13 languages | Global developer audience | `Localize/loc/` coverage |
| 6 | **Test coverage** of every shipped capability | Prevents regression | `.dev/features.json` × CI matrix |

## Stakeholders

| Role | Concerns |
|------|----------|
| Developer (P1–P3) | Time-to-first-run, error clarity, language support |
| CI/DevOps (P4) | Determinism, headless mode, pipeline templates |
| IT admin (P5) | Manifest validation, permission visibility |
| Microsoft 365 platform teams | Schema versions, telemetry of capability usage |
| Microsoft Foundry / Copilot teams | DA / API plugin / MCP feature parity |
| OSS contributors (P6) | Discoverability of conventions, fast feedback loop |
