# Infrastructure

Engine-neutral infrastructure documents for the Microsoft 365 Agents Toolkit.
This folder answers a single question:

> *"What runtime substrate does the toolkit (or what it scaffolds) live on, and
> what cross-cutting contracts govern provisioning, identity, deployment, and
> operations?"*

Infrastructure is engine-external: it describes the cloud / identity / pipeline
substrate the toolkit produces, consumes, or runs against. The internal shape of
the engine itself lives in [`docs/02-architecture/`](../02-architecture/README.md).

## What lives here

Three kinds of content, in this order:

1. **Chosen substrate and topology** — target clouds and runtimes, resource-group
   layout, region and environment model (`dev`, `staging`, `prod`), tagging,
   naming.
2. **Provisioning / runtime contracts** — Bicep / Terraform module shape,
   parameter and output contracts, identity baselines (managed identity,
   federated credentials, RBAC), secret handling, network assumptions, telemetry
   backend, CI/CD and release-channel conventions.
3. **Resulting constraints** — rules every starter template, driver, or
   deployment plan must follow because of (1) and (2). Constraints are the
   *output* of these decisions, not the definition of infrastructure.

Binding infrastructure decisions that are hard to reverse or cross-cutting are
recorded as ADRs under [`docs/02-architecture/`](../02-architecture/README.md);
pattern pages in this folder link back to those ADRs rather than restating them.

## What does NOT live here

- Engine-internal driver implementations or operation specs → [`docs/04-specs/`](../04-specs/README.md).
- How the engine is internally composed → [`docs/02-architecture/`](../02-architecture/README.md).
- Concrete starter infrastructure that ships *to users* under `templates/` —
  those are shipping artifacts, not the doc-level contract. This folder may
  reference them as examples but does not own them.
- Product intent or scenario flows → [`docs/01-product/`](../01-product/README.md).

## Status

This folder is being populated. Add infrastructure pages here as patterns
stabilize across templates and supported clouds.
