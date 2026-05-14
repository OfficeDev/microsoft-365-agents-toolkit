# 03 — Infrastructure

What the toolkit **provisions on behalf of users**. (For the toolkit's *own* runtime — short version: it has none, it's a local tool — see [02-architecture/07-deployment-view.md](../02-architecture/07-deployment-view.md).)

This section explains the Azure / Entra / Bot-Framework / Microsoft-365 topology the lifecycle drivers create, and the conventions the shipped Bicep templates follow.

## Pages

- [overview.md](overview.md) — the topology pattern catalogue
- [azure-topology.md](azure-topology.md) — App Service, Functions, Static Web Apps shapes
- [bicep-templates.md](bicep-templates.md) — what ships in each `infra/` folder
- [networking-and-identity.md](networking-and-identity.md) — Entra app reg, managed identity, role assignments
- [secrets-and-keyvault.md](secrets-and-keyvault.md) — secret handling patterns
- [cicd-pipelines.md](cicd-pipelines.md) — GitHub Actions / ADO templates emitted to user repos
- [local-dev-environment.md](local-dev-environment.md) — dev tunnels, M365 Agents Playground
