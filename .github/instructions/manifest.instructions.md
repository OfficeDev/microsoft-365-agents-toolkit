---
description: Instructions for @microsoft/app-manifest package - TypeScript type definitions and converters for Microsoft 365 App manifests
applyTo: 'packages/manifest/**'
---

# Package: `@microsoft/app-manifest` (packages/manifest)

## Overview
This package provides **TypeScript type definitions and converters** for Microsoft 365 App manifests. It automatically generates strongly-typed interfaces from official Microsoft JSON schemas and provides utilities for manifest manipulation, validation, and I/O.

## How It Works

### Schema-Driven Code Generation
1. **Schema Collection**: JSON schemas are stored in `src/json-schemas/` organized by manifest type and version
2. **Code Generation**: Run `npm run convert` to execute `convert.js`, which uses `quicktype-core-jayzhang` to generate TypeScript types from JSON schemas
3. **Generated Types**: Output goes to `src/generated-types/` with type interfaces and `Convert` classes for each version
4. **Discriminated Unions**: All versions are combined into union types (`TeamsManifest`, `DeclarativeAgentManifest`, `APIPluginManifest`)

### Directory Structure
```
src/
├── json-schemas/                    # Source JSON schemas (SOURCE OF TRUTH for versions)
│   ├── teams/                       # Teams manifest schemas (check subfolders for versions)
│   └── copilot/
│       ├── declarative-agent/       # Declarative agent schemas (check subfolders for versions)
│       └── plugin/                  # API plugin schemas (check subfolders for versions)
├── generated-types/                 # Auto-generated TypeScript types
│   ├── index.ts                     # Converters, utilities, union types (defines ConverterMaps)
│   ├── teams/                       # TeamsManifestV{version}.ts files
│   └── copilot/
│       ├── declarative-agent/       # DeclarativeAgentManifestV{version}.ts files
│       └── plugin/                  # ApiPluginManifestV{version}.ts files
├── wrappers/                        # OOP wrappers with fluent APIs (check index.ts for exports)
│   ├── index.ts                     # Exports all wrappers and related types/constants
│   ├── BaseManifest.ts              # Abstract base class with common functionality
│   ├── PluginManifestWrapper.ts     # Wrapper for API Plugin manifests
│   └── DeclarativeAgentManifestWrapper.ts  # Wrapper for Declarative Agent manifests
└── index.ts                         # Main entry point, re-exports
```

## Supported Manifests

> **To discover current supported versions**: List the subfolders in `src/json-schemas/` directories, or check the union types and `ConverterMap` objects in `src/generated-types/index.ts`.

| Manifest Type | Version Discovery Location | Version Field |
|---------------|---------------------------|---------------|
| **Teams Manifest** | `src/json-schemas/teams/` subfolders | `manifestVersion` |
| **Declarative Agent** | `src/json-schemas/copilot/declarative-agent/` subfolders | `version` |
| **API Plugin** | `src/json-schemas/copilot/plugin/` subfolders | `schema_version` |

## Key APIs

### Union Types
> **Check `src/generated-types/index.ts`** for the current union type definitions and all supported versions.

```typescript
// Pattern: Union of all version-specific types
type TeamsManifest = TeamsManifestV1D0 | TeamsManifestV1D1 | ... ;
type DeclarativeAgentManifest = DeclarativeAgentManifestV1D0 | ... ;
type APIPluginManifest = APIPluginManifestV2D1 | ... ;
```

## Adding a New Schema Version
1. Add the new schema JSON file to the appropriate folder in `src/json-schemas/`
2. Run `npm run convert` to regenerate types
3. Update `src/generated-types/index.ts` to:
   - Import the new version module
   - Add to the union type
   - Add converter mapping in the appropriate `ConverterMap`

## OOP Manifest Wrappers

For convenient manifest manipulation, the package provides **OOP wrappers** with fluent APIs in `src/wrappers/`:

### Available Wrappers

> **To discover available wrappers and their APIs**: Check `src/wrappers/index.ts` for exports, then read the corresponding wrapper files.

| Wrapper Class | Purpose | Location |
|---------------|---------|----------|
| `BaseManifest<T>` | Abstract base class with common functionality | `src/wrappers/BaseManifest.ts` |
| `PluginManifestWrapper` | Wrapper for API Plugin manifests | `src/wrappers/PluginManifestWrapper.ts` |
| `DeclarativeAgentManifestWrapper` | Wrapper for Declarative Agent manifests | `src/wrappers/DeclarativeAgentManifestWrapper.ts` |

### Design Patterns

All wrappers follow consistent patterns. Read the wrapper source file to discover all available methods.

#### 1. Static Factory Methods
Every wrapper provides these factory methods:
- `read(filePath)` - Async read from file
- `readSync(filePath)` - Sync read from file  
- `fromJSON(json)` - Create from JSON string
- `create(init)` - Create new with required fields

#### 2. Getters (Read-only Access)
Property getters provide read-only access to manifest data. Look for `get propertyName()` in the wrapper.

#### 3. Setters (Fluent API)
Methods prefixed with `set` modify scalar properties and return `this` for chaining.

#### 4. Collection Operations (Fluent API)
For array properties (functions, runtimes, actions, capabilities, etc.), wrappers provide:
- `add{Item}(...)` - Add an item, returns `this`
- `remove{Item}(...)` - Remove by identifier, returns `this`
- `has{Item}(...)` - Check existence, returns `boolean`
- `get{Item}(...)` - Get single item, returns item or `undefined`

#### 5. Base Class Methods (Inherited)
From `BaseManifest<T>`:
- `data` - Read-only access to raw manifest data
- `filePath` - Original file path (if loaded from file)
- `isDirty` - Whether manifest has unsaved changes
- `save(filePath?)` - Save to file
- `validate()` - Validate against JSON schema (abstract, implemented by subclass)
- `toJSON()` - Serialize to JSON string (abstract, implemented by subclass)
- `clone()` - Deep clone (abstract, implemented by subclass)

### Schema-Derived Types and Constants

Wrappers use types derived from the latest JSON schema to ensure type safety:

```typescript
// Types are derived from the latest manifest version
type LatestManifestType = APIPluginManifestLatest;  // or DeclarativeAgentManifestLatest
type RuntimeObject = NonNullable<LatestManifestType["runtimes"]>[number];

// Enums are auto-generated as const objects matching schema values
export const RuntimeType: { readonly [K in RuntimeTypeValue]: K } = {
  OpenApi: "OpenApi",
  LocalPlugin: "LocalPlugin",
  // ... values from schema
} as const;
```

> **To discover available enum values**: Check the `const` exports in the wrapper file (e.g., `RuntimeType`, `CapabilityName`). These are derived from the JSON schema and reflect all valid values.

### Usage Pattern
```typescript
import { PluginManifestWrapper, DeclarativeAgentManifestWrapper } from "@microsoft/app-manifest";

// Read -> Modify (fluent chain) -> Save
const manifest = await SomeManifestWrapper.read("manifest.json");
manifest
  .setSomeProperty("value")
  .addSomeItem(...);
await manifest.save();
```

### Extending Wrappers
To add a new wrapper:
1. Create a new class extending `BaseManifest<T>` in `src/wrappers/`
2. Define internal types derived from the latest manifest version (see existing wrappers)
3. Implement required abstract methods: `validate()`, `toJSON()`, `clone()`
4. Add static factory methods following the pattern
5. Add getters, setters, and collection operations as needed
6. Export any schema-derived const enums from the wrapper
7. Export from `src/wrappers/index.ts`
