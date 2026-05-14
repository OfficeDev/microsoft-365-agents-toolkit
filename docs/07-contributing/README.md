# 07 — Contributing

Playbooks for the most common contributor tasks. The repo-root [`CONTRIBUTING.md`](../../CONTRIBUTING.md) covers the legal / CLA / Code of Conduct side; this section is the **how** of common engineering changes.

## Pages

- [getting-started.md](getting-started.md) — first-time setup
- [dev-setup.md](dev-setup.md) — toolchain and commands
- [adding-a-template.md](adding-a-template.md)
- [adding-a-driver.md](adding-a-driver.md)
- [adding-a-cli-command.md](adding-a-cli-command.md)
- [adding-a-feature-flag.md](adding-a-feature-flag.md)
- [docs-style-guide.md](docs-style-guide.md)

## Cross-cutting rules

- Conventional Commits (commitlint enforces).
- Pull requests against `dev` (default branch).
- `pnpm install` to sync workspaces.
- Tests required for every feature / fix.
- Lint must be 0 errors; warnings allowed for `no-explicit-any`.
