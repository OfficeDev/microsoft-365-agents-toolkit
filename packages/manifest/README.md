## 📦 Package Name

`@microsoft/app-manifest`

---

## ✨ Features

A collection of TypeScript definitions and converters for Microsoft 365 App manifests.

- **Strongly‑typed interfaces** for three manifest types with all versions generated from Microsoft’s official [JSON schemas](https://developer.microsoft.com/json-schemas/), guaranteeing compile‑time type correctness:  
  - Teams Manifest: 1.0, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 1.10, 1.11, 1.12, 1.13, 1.14, 1.15, 1.16, 1.17, 1.19, 1.20, 1.21, 1.22, devPreview  
  - Declarative Agent Manifest: v1.0, v1.2, v1.3, v1.4  
  - API Plugin Manifest: v2.1, v2.2

- **Bi‑directional conversions** from JSON to manifest or from manifest to JSON(`jsonToManifest` and `manifestToJson`) with runtime type validation.

- **Schema validation utilities** to validating manifests against schemas  

- **File I/O helpers** to conveniently load manifest from file or dump manifest to file in JSON format.

- **Modular versioning**: manifests organized per version file, avoiding type collisions.

## 📥 Installation

```bash
npm install @microsoft/app-manifest 
```

## 📖 Usage

### Manipulate manifest

**Automatic version inference**

You can assign a manifest object directly to the discriminated union types (`TeamsManifest`, `DeclarativeAgentManifest` or `APIPluginManifest`) without specifying a concrete version type. TypeScript will infer the specific version based on the `manifestVersion` field:

![not_specify_version](https://github.com/user-attachments/assets/2f10873b-974a-4998-a80c-6bdd6609bdfc)

**Manually specified version**

You can specify a concrete version type:

![specify_version](https://github.com/user-attachments/assets/d6a566af-8b68-41cf-b1a2-a1083e16e6c7)

### Manifest to/from JSON converters

Convert JSON string to Teams manifest type and check the version at run time:

```typescript
const json = "{ \"manifestVersion\": \"1.20\", \"id\": \"app-id\", ...}";
const manifest = TeamsManifestConverter.jsonToManifest(json);
if (manifest.manifestVersion === "1.20") {
  // TypeScript will infer the type as TeamsManifestV1D20
  const manifestV1D20 = manifest as TeamsManifestV1D20.TeamsManifestV1D20;
  // You can now access properties specific to TeamsManifestV1D20
}
```

Convert JSON string to Teams manifest type by specifying the version at compile time:

```typescript
const json = "{ \"manifestVersion\": \"1.20\", \"id\": \"app-id\", ...}";
const manifest = TeamsManifestConverter.jsonToManifest(json) as TeamsManifestV1D20.TeamsManifestV1D20;
// You can now access properties specific to TeamsManifestV1D20
```

Convert JSON string to declarative agent manifest type and api plugin manifest type:

```typescript
const daManifest = DeclarativeAgentManifestConverter.jsonToManifest(daManifestJSON) as DeclarativeAgentManifest;
const pluginManifest = ApiPluginManifestConverter.jsonToManifest(pluginManifestJSON) as APIPluginManifest;
```

Convert manifest object to JSON string:

```typescript
const jsonString = TeamsManifestConverter.manifestToJson(manifest);
const daJsonString = DeclarativeAgentManifestConverter.manifestToJson(daManifest);
const pluginJsonString = ApiPluginManifestConverter.manifestToJson(pluginManifest);
```

Note that the converts to/from JSON will throw runtime type check failures.

### Manifest utilities

Validate manifest against schema: 

```typescript
const failures = await AppManifestUtils.validateAgainstSchema(manifest);
```

Read and write manifest:

```typescript
const teamsManifestPath = "path/to/your/teams_manifest.json"; 
// read Teams manifest with type check
const teamsManifest1 = await AppManifestUtils.readTeamsManifest(teamsManifestPath);
// read Teams manifest and validate against schema
const [teamsManifest2, failedValidations1] = await AppManifestUtils.readAndValidateTeamsManifest(teamsManifestPath);
 
const daManifestPath = "path/to/your/da_manifest.json"; 
// read declarative agent manifest with type check
const daManifest1 = await AppManifestUtils.readDeclarativeAgentManifest(daManifestPath);
// read declarative agent manifest and validate against schema
const [daManifest2, failedValidations2] = await AppManifestUtils.readAndValidateDeclarativeAgentManifest(daManifestPath);

const pluginManifestPath = "path/to/your/plugin_manifest.json"; 
// read API plugin manifest with type check
const pluginManifest1 = await AppManifestUtils.readApiPluginManifest(pluginManifestPath);
// read API plugin manifest and validate against schema
const [pluginManifest2, failedValidations3] = await AppManifestUtils.readAndValidateApiPluginManifest(pluginManifestPath);
```