---
description: Instructions for @microsoft/app-manifest package - TypeScript type definitions and converters for Microsoft 365 App manifests
applyTo: 'packages/manifest/**,packages/**manifest*,packages/**Manifest*'
---

# Package: `@microsoft/app-manifest` (packages/manifest)

## Overview

This package provides **TypeScript type definitions, converters, and OOP wrappers** for Microsoft 365 App manifests. It automatically generates strongly-typed interfaces from official Microsoft JSON schemas and provides utilities for manifest manipulation, validation, and I/O.

## Three Manifest Types

| Manifest Type | Purpose | Version Field | Latest Type Alias |
|---------------|---------|---------------|-------------------|
| **Teams Manifest** | Core app manifest for M365 apps (bots, tabs, extensions) | `manifestVersion` | `TeamsManifestLatest` |
| **Declarative Agent Manifest** | AI agent instructions, knowledge sources, actions | `version` | `DeclarativeAgentManifestLatest` |
| **API Plugin Manifest** | Plugin capabilities, REST API operations | `schema_version` | `APIPluginManifestLatest` |

> **Note**: Check `src/json-schemas/` for available versions. The `*Latest` type aliases always point to the most recent stable version.

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
│   ├── teams/                       # Teams manifest schemas (v1.0-v1.24, devPreview)
│   └── copilot/
│       ├── declarative-agent/       # Declarative agent schemas (v1.0-v1.6)
│       └── plugin/                  # API plugin schemas (v2.1-v2.4)
├── generated-types/                 # Auto-generated TypeScript types
│   ├── index.ts                     # Union types, converters, AppManifestUtils
│   ├── teams/                       # TeamsManifestV{version}.ts files
│   └── copilot/
│       ├── declarative-agent/       # DeclarativeAgentManifestV{version}.ts files
│       └── plugin/                  # ApiPluginManifestV{version}.ts files
├── wrappers/                        # OOP wrappers with fluent APIs
│   ├── index.ts                     # Exports all wrappers and related types/constants
│   ├── BaseManifest.ts              # Abstract base class with common functionality
│   ├── TeamsManifestWrapper.ts      # Wrapper for Teams manifests
│   ├── PluginManifestWrapper.ts     # Wrapper for API Plugin manifests
│   └── DeclarativeAgentManifestWrapper.ts  # Wrapper for Declarative Agent manifests
├── manifest.ts                      # @deprecated - Legacy Teams manifest types
├── declarativeCopilotManifest.ts    # @deprecated - Legacy declarative agent types
├── pluginManifest.ts                # @deprecated - Legacy plugin types
└── index.ts                         # Main entry point, re-exports
```

## Key APIs

### Union Types (from `generated-types/index.ts`)
```typescript
// Union of all supported versions for each manifest type
type TeamsManifest = TeamsManifestV1D0 | TeamsManifestV1D1 | ... | TeamsManifestVDevPreview;
type DeclarativeAgentManifest = DeclarativeAgentManifestV1D0 | ...;
type APIPluginManifest = APIPluginManifestV2D1 | ...;
type AppManifest = TeamsManifest | DeclarativeAgentManifest | APIPluginManifest;
```

### Converters (from `generated-types/index.ts`)
```typescript
// Parse JSON to typed manifest
TeamsManifestConverter.jsonToManifest(json: string): TeamsManifest
DeclarativeAgentManifestConverter.jsonToManifest(json: string): DeclarativeAgentManifest
ApiPluginManifestConverter.jsonToManifest(json: string): APIPluginManifest

// Serialize manifest to JSON
TeamsManifestConverter.manifestToJson(manifest: TeamsManifest): string
DeclarativeAgentManifestConverter.manifestToJson(manifest: DeclarativeAgentManifest): string
ApiPluginManifestConverter.manifestToJson(manifest: APIPluginManifest): string
```

### AppManifestUtils (from `generated-types/index.ts`)
```typescript
// Read manifests from file
AppManifestUtils.readTeamsManifest(filePath): Promise<TeamsManifest>
AppManifestUtils.readDeclarativeAgentManifest(filePath): Promise<DeclarativeAgentManifest>
AppManifestUtils.readApiPluginManifest(filePath): Promise<APIPluginManifest>

// Read with validation
AppManifestUtils.readAndValidateTeamsManifest(filePath): Promise<[TeamsManifest, string[]]>
AppManifestUtils.readAndValidateDeclarativeAgentManifest(filePath): Promise<[DeclarativeAgentManifest, string[]]>
AppManifestUtils.readAndValidateApiPluginManifest(filePath): Promise<[APIPluginManifest, string[]]>

// Write manifests
AppManifestUtils.writeTeamsManifest(filePath, manifest): Promise<void>
AppManifestUtils.writeDeclarativeAgentManifest(filePath, manifest): Promise<void>
AppManifestUtils.writeApiPluginManifest(filePath, manifest): Promise<void>

// Validation
AppManifestUtils.validateAgainstSchema(manifest, schema?): Promise<string[]>
AppManifestUtils.fetchSchema(schemaUrl): Promise<JSONSchemaType<AppManifest>>
```

## OOP Manifest Wrappers

For convenient manifest manipulation, the package provides **OOP wrappers** with fluent APIs in `src/wrappers/`:

### Available Wrappers

| Wrapper Class | Purpose | Exported Enums |
|---------------|---------|----------------|
| `TeamsManifestWrapper` | Teams app manifests | `DefaultInstallScope`, `StaticTabScope`, `ConfigurableTabScope`, `ComposeExtensionTypeEnum` |
| `DeclarativeAgentManifestWrapper` | Declarative agent manifests | `CapabilityName` |
| `PluginManifestWrapper` | API plugin manifests | `RuntimeType` |

### Design Patterns

All wrappers follow consistent patterns:

#### 1. Static Factory Methods
```typescript
SomeWrapper.read(filePath)     // Async read from file
SomeWrapper.readSync(filePath) // Sync read from file
SomeWrapper.fromJSON(json)     // Create from JSON string
SomeWrapper.create(init)       // Create new with required fields
```

#### 2. Fluent Setters (return `this` for chaining)
```typescript
manifest
  .setName("My App")
  .setDescription("Short", "Full")
  .setVersion("1.0.0");
```

#### 3. Collection Operations
```typescript
manifest.addBot(botId, scopes)      // Add item
manifest.removeBot(botId)           // Remove by identifier
manifest.hasBot(botId)              // Check existence
manifest.getBot(botId)              // Get single item
```

#### 4. Common Properties
```typescript
manifest.data        // Read-only raw manifest data
manifest.filePath    // Original file path (if loaded from file)
manifest.isDirty     // Whether manifest has unsaved changes
manifest.save()      // Save to file
manifest.validate()  // Validate against JSON schema
manifest.toJSON()    // Serialize to JSON string
manifest.clone()     // Deep clone
```

### Usage Examples

```typescript
import {
  TeamsManifestWrapper,
  DeclarativeAgentManifestWrapper,
  PluginManifestWrapper,
  DefaultInstallScope,
  CapabilityName,
} from "@microsoft/app-manifest";

// Teams Manifest
const teamsManifest = await TeamsManifestWrapper.read("manifest.json");
teamsManifest
  .setName("My Teams App")
  .setDefaultInstallScope(DefaultInstallScope.personal)
  .addBot("bot-id", ["personal", "team"])
  .addStaticTab("home", "Home", "https://example.com/tab")
  .addDeclarativeAgent("agent1", "declarativeAgent.json");
await teamsManifest.save();

// Declarative Agent Manifest
const agentManifest = await DeclarativeAgentManifestWrapper.read("agent.json");
agentManifest
  .setName("My Agent")
  .setInstructions("You are a helpful assistant.")
  .addCapability(CapabilityName.WebSearch);
await agentManifest.save();

// API Plugin Manifest
const pluginManifest = await PluginManifestWrapper.read("plugin.json");
pluginManifest
  .setNameForHuman("My Plugin")
  .addFunction({ name: "getData", description: "Gets data" });
await pluginManifest.save();
```

## Adding a New Schema Version

1. Sync schemas from upstream first (preferred): run `node download.js` in `packages/manifest` so new schema folders are pulled from `microsoft/json-schemas`
2. If needed, add/update schema JSON files in the appropriate folder under `src/json-schemas/`
3. Run `npm run convert` to regenerate types
4. Update `src/generated-types/index.ts` to:
   - Import the new version module
   - Add to the union type
   - Add converter mapping in the appropriate `ConverterMap`
   - Re-export any newly introduced enums/types that callers consume (for example, schema-added enums)
5. If this is a new "latest" version, update the `*Latest` type alias
6. Ensure the new schema version is registered in converter maps (for example `daConverterMap` / `TeamsManifestConverterMap`) to avoid fallback unchecked casts that skip validation
7. Verify map/schema parity with `npx mocha test/converterMapParity.test.ts`

## Extending Wrappers

To add a new wrapper:
1. Create a new class in `src/wrappers/` (extend `BaseManifest<T>` if the manifest type satisfies `Record<string, unknown>`)
2. Define internal types derived from the latest manifest version
3. Implement required methods: `validate()`, `toJSON()`, `clone()`
4. Add static factory methods: `read()`, `readSync()`, `fromJSON()`, `create()`
5. Add getters, setters, and collection operations as needed
6. Export any schema-derived const enums from the wrapper
7. Export from `src/wrappers/index.ts`

## Deprecated Types

The following manually-written type definitions are **deprecated** and should be replaced with auto-generated types:

| Deprecated | Replacement |
|------------|-------------|
| `TeamsAppManifest` (manifest.ts) | `TeamsManifest` or `TeamsManifestLatest` |
| `DeclarativeCopilotManifestSchema` (declarativeCopilotManifest.ts) | `DeclarativeAgentManifest` or `DeclarativeAgentManifestLatest` |
| `PluginManifestSchema` (pluginManifest.ts) | `APIPluginManifest` or `APIPluginManifestLatest` |
| `ManifestUtil` (index.ts) | `AppManifestUtils` |

These legacy types remain for backward compatibility but will be removed in a future major version.
