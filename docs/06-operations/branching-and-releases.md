# Branching and releases

## Branches

| Branch | Purpose |
|--------|---------|
| `dev` | Default integration branch. PRs merge here. |
| `main` | Release branch. Tagged for stable releases. |
| `release/*` | Hot-fix lines for shipped versions. |
| `users/*` | Personal feature branches. Deleted after merge. |

## PR lifecycle

1. Open PR against `dev`.
2. CodeQL + CI must be green.
3. Reviewer approval (CODEOWNERS auto-suggests).
4. Squash merge with a Conventional Commit subject.
5. Renovate / Dependabot PRs get the same treatment.

## Release cadence

- VS Code extension: weekly prerelease (Insiders) + monthly stable.
- CLI v3 + fx-core: cut from `main` per stable release.
- v4 packages (`core-next`, `cli-next`): preview releases on `dev`; not published to `latest` tag yet.

## Hotfixes

1. Cherry-pick to `release/x.y.z`.
2. Run full CI on the release branch.
3. Tag and publish.

## Tags

Git tags follow `vMAJOR.MINOR.PATCH` per package. Lerna publishes via the `npm publish` flow with provenance.
