# ADR-0023 - Create input policy ownership

- **Status:** Accepted (user-delegated design decision, 2026-09-09).
- **Supersedes:** ADR-0016 decision 1's closed descriptor field set only; all other decisions remain binding.
- **Scope:** v4 create-input policy placement; no new product flow.

## Context

Create-input composition currently decides Python Preview presentation from template
IDs, implements C# availability, and repairs a search-specific OpenAPI answer after
the shared walk. These decisions obscure the ownership promised by ADR-0016.

## Decision

1. Add optional `descriptor.languageOptions`, an array of presentation overrides
   with `id`, `label`, `description`, and `keyPrefix`. IDs are unique and must
   belong to `descriptor.languages`. Overrides cannot add languages, carry
   configuration payloads, or execute conditions. The existing localization
   mechanism supports both literal fallback strings and NLS keys.
2. A named `create.languages` provider owns default labels and the existing
   C# surface/feature-flag policy. It receives descriptor data and injected
   surface/flags, never a template-ID allowlist. The common floor only decides
   cardinality, default selection, and question placement from resolved options.
3. The presentation field and provider are introduced at engine capability
   version `6.12.0`. Packages using the field must declare at least that version;
   older engines reject them through the existing reverse compatibility gate.
   Packages without the field remain valid and receive default presentation,
   without template-ID-based Preview inference.
4. OpenAPI providers own canonical source derivation. The operation provider
   declares `apiSpecLocation` in its derived schema; template bindings consume
   `derived.openapi.operations.apiSpecLocation`. The engine does not copy an
   OpenAPI search answer into a different top-level answer after the walk.
   Search question names and CLI input names remain unchanged. The synthetic
   top-level search alias is not part of the new internal answer contract.
   Consuming this derived output requires engine `6.12.0`; the existing
   operation-listing capability keeps its original version floor. A named
   compatibility adapter preserves the pre-6.12 DA search package's render
   binding only when its old binding and derived source are present. It does
   not modify collected answers, new packages, or explicit source bindings.
5. Keep the shared question walk, staged package loading, existing language
   availability, prompt ordering, and generated current-template artifacts.
   Do not introduce an answer-alias DSL, arbitrary hooks, or a second walk.
   The legacy binding adapter is a source-owned compatibility migration,
   not a template-defined postprocessor.

## Alternatives

- Moving the template-ID allowlist into another utility preserves hidden
  template coupling; presentation metadata is independently testable.
- Putting a complete language question into every template duplicates common
  floor semantics and risks changing Back and singleton behavior.
- Adding generic post-walk business hooks creates another extension contract
  where the existing provider-derived output contract is sufficient.

## Consequences

The descriptor schema and semantic validation jointly protect presentation data.
Runtime composition is independent of template IDs and OpenAPI field names.
Provider and package tests own policy; floor tests own only composition.

The new shape is specified and tested by CLEAN-01 through CLEAN-07 in
[collect-create-inputs](../../03-specs/operations/scaffolding/collect-create-inputs.md).