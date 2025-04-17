## 📦 Package Name

`@microsoft/app-manifest`

---

## 🚀 Summary

A collection of TypeScript definitions and converters for Microsoft 365 App manifests, including:

- **Strongly‑typed interfaces** for three manifest types with all versions:  
  - Teams Manifest: 1.0, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 1.10, 1.11, 1.12, 1.13, 1.14, 1.15, 1.16, 1.17, 1.19, 1.20, devPreview
  - Declarative Agent Manifest: v1.0, v1.2, v1.3
  - API Plugin Manifest: v2.1, v2.2
- **Conversion utilities** between JSON strings and typed manifest objects.  
- **Utility tools** for:  
  - Validating manifests against manifest schemas  
  - Reading/writing manifest files from disk with type checking  

---

## ✨ Features

- **Type‑safe definitions** of all versions generated from Microsoft’s official [JSON schemas](https://developer.microsoft.com/json-schemas/), guaranteeing compile‑time type correctness.

- **Bi‑directional conversion functions** (`jsonToManifest` and `manifestToJson`) for all types and all versions of manifest with runtime type validation.

- **Schema validation utilities** to generate validation errors.

- **File I/O helpers** to conveniently load and dump manifest files in JSON format.

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

Convert JSON string to manifest type and check the version at run time:

```typescript
const json = "{ \"manifestVersion\": \"1.20\", \"id\": \"app-id\", ...}";
const manifest = TeamsManifestConverter.jsonToManifest(json);
if (manifest.manifestVersion === "1.20") {
  // TypeScript will infer the type as TeamsManifestV1D20
  const manifestV1D20 = manifest as TeamsManifestV1D20;
  // You can now access properties specific to TeamsManifestV1D20
}
```

Convert JSON string to manifest type by specifying the version at compile time:

```typescript
const json = "{ \"manifestVersion\": \"1.20\", \"id\": \"app-id\", ...}";
const manifest = TeamsManifestConverter.jsonToManifest(json) as TeamsManifestV1D20;
// You can now access properties specific to TeamsManifestV1D20
```

Convert manifest object to JSON string:

```typescript
const jsonString = TeamsManifestConverter.manifestToJson(manifest);
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
// read declarative agent manifest with type check
const pluginManifest1 = await AppManifestUtils.readApiPluginManifest(pluginManifestPath);
// read declarative agent manifest and validate against schema
const [pluginManifest2, failedValidations3] = await AppManifestUtils.readAndValidateApiPluginManifest(pluginManifestPath);
```