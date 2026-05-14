# 04 — Specs: Data model

The structures the toolkit operates on. Includes both **on-disk artefacts** in user projects and **in-memory contracts** consumed by the engine.

## Pages

- [overview.md](overview.md) — entity-relationship view
- [storage-locations.md](storage-locations.md) — where every persistent thing lives
- [entities/](entities/README.md) — one page per entity
- [schemas/](schemas/README.md) — JSON / YAML schema references

For TypeScript types backing these entities, see the source under `packages/api`, `packages/manifest`, `packages/core-next/src/api`, and the per-package reference pages.
