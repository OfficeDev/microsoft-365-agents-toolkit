# 01 — Product (PRD)

This is the **product layer**: who we serve, what they need, what we promise, how we measure success. It feeds the architecture below it. Think of it as the "*why*" behind every other section.

## Pages

- [prd-overview.md](prd-overview.md) — product requirements summary
- [personas.md](personas.md) — developer / IT-admin personas
- [capabilities-matrix.md](capabilities-matrix.md) — what is buildable today, mapped to packages
- [v3-feature-inventory.md](v3-feature-inventory.md) — deep enumeration of v3-shipped capabilities (extracted from source; allowed as v4 design input)
- [success-metrics.md](success-metrics.md) — how we measure that we're delivering
- [scenarios/](scenarios/README.md) — end-to-end user scenarios

## Authoritative source

The machine-readable feature inventory is [`.dev/features.json`](../../.dev/features.json), wrapped by [`packages/fx-core/tests/integration/featureRegistry.ts`](../../packages/fx-core/tests/integration/featureRegistry.ts). Every page in this section that lists capabilities should derive from that file, not duplicate it.
