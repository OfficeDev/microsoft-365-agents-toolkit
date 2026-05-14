# v3 `QuestionNames` enum

> **FORBIDDEN AS v4 DESIGN INPUT.** See [`../README.md`](../README.md).
>
> **Exception:** v4 may reuse individual question *concepts* (e.g. "AppName", "Folder") because they map to product-level UX requirements. v4 must **not** reuse the full enum, the names verbatim, or the v3 question-tree structure — those carry v3's accidental complexity (e.g. typos like `Scratch`/`SctatchYes`, `RepalceTabUrl`).

Source: [`packages/fx-core/src/question/questionNames.ts`](../../../packages/fx-core/src/question/questionNames.ts). **147 enum members** at extraction time — the source file is the truth.

## Full enum

```typescript
enum QuestionNames {
  // Project shape
  Scratch
  SctatchYes                 // sic
  AppName
  Folder
  ProjectPath
  ProgrammingLanguage
  ProjectType
  Capabilities
  TeamsAppType
  TeamsCapability
  BotTrigger
  Runtime
  SafeProjectName
  SolutionName

  // SPFx
  SPFxSolution
  SPFxInstallPackage
  SPFxFramework
  SPFxWebpartName
  SPFxWebpartDesc
  SPFxFolder

  // Office Add-in
  OfficeAddinFolder
  OfficeAddinManifest
  OfficeAddinTemplate
  OfficeAddinHost
  OfficeAddinImport
  OfficeAddinFramework

  // Samples / migration
  Samples
  ReplaceContentUrl
  ReplaceWebsiteUrl
  ReplaceBotIds
  RepalceTabUrl              // sic
  ExistingTabEndpoint

  // OpenAPI / API plugins
  FromExistingApi
  SearchOpenAPISpecQuery
  SelectOpenApiSpec
  OpenAPISpecType
  ApiSpecLocation
  ApiOperation
  ActionManifestPath
  MeArchitectureType
  ApiSpecApiKey
  ApiSpecApiKeyConfirm
  ApiAuth
  OauthClientSecret
  OauthClientId
  OauthConfirm
  CustomCopilotRag
  CustomCopilotAssistant

  // LLM provider
  LLMService
  OpenAIKey
  OpenAIEmbeddingModel
  OpenAIAssistantID
  AzureOpenAIKey
  AzureOpenAIEndpoint
  AzureOpenAIDeploymentName
  AzureOpenAIEmbeddingDeploymentName
  AzureOpenAIAssistantId
  AzureAISearchApiKey
  AzureAISearchEndpoint

  // Foundry
  FoundryEndpoint
  FoundryAgentId

  // Misc project lifecycle
  Features
  Env
  SourceEnvName
  TargetEnvName
  TargetResourceGroupName
  NewResourceGroupName
  NewResourceGroupLocation
  NewTargetEnvName

  // Manifest paths
  TeamsAppManifestFilePath
  LocalTeamsAppManifestFilePath
  AadAppManifestFilePath
  TeamsAppPackageFilePath
  ConfirmManifest
  ConfirmLocalManifest
  ConfirmAadManifest
  OutputZipPathParamName
  OutputManifestParamName
  OutputFolderParamName
  ValidateMethod
  AppPackagePath

  // Sideloading / install
  M365Host
  ManifestPath
  ManifestId
  TeamsAppId
  TitleId
  UserEmail
  UninstallMode
  UninstallModeManifestId
  UninstallModeEnv
  UninstallModeTitleId
  UninstallOptions
  UninstallOptionM365
  UninstallOptionTDP
  UninstallOptionBot

  // Collaboration / sharing
  collaborationAppType
  ShareOperation
  ShareScope
  RemoveUsers
  DestinationApiSpecFilePath
  SyncManifest

  // Action / plugin extension
  ActionType
  WithPlugin
  ImportPlugin
  PluginManifestFilePath
  NewPluginManifestFileName
  PluginOpenApiSpecFilePath
  KnowledgeSource
  OneDriveSharePointURL
  OneDriveSharePointContent
  WebContent
  SearchType
  GCContent
  GCList
  GCInput
  GCName
  GCConnectionId
  AuthName
  TemplateName
  EmbeddedKnowledgeFiles

  // OAuth detail
  OAuthAuthorizationUrl
  OAuthTokenUrl
  OAuthRefreshUrl
  OAuthScope
  OauthPKCE
  ApiKeyIn
  ApiKeyName

  // TypeSpec
  TypeSpecProjectType
  DeclarativeAgentManifestPath

  // Misc DA
  SensitivityLabel
  SelectPluginManifest
  SelectOpenAPISpecFromPlugin
  SelectPluginId
  DAMetaOSCapability
  MCPServerType

  // MCP server (for-DA action) — added after the original extraction
  MCPLocalServer
  MCPLocalServerIdentifier
  MCPForDAServerUrl
  MCPForDAServerName
  MCPForDATool
  MCPForDAAvailableTools
  MCPForDAPreFetchTools
  MCPForDAAuth
  MCPForDAAuthMetadataUrl
  MCPForDAAuthWellKnownUrl
  MCPForDAAuthType
  MCPToolsFilePath
}
```

## Notes

- **147 entries total.** Several typos (`SctatchYes` for `ScratchYes`, `RepalceTabUrl` for `ReplaceTabUrl`) are baked in for back-compat — verified to exist in source.
- Mixed casing styles: `collaborationAppType` (camelCase) vs the rest (PascalCase).
- Many entries are scoped to a single template branch (e.g. `GCName`, `GCList`, `GCInput` are all graph-connector-only).
- The MCP-DA family (12 entries) was added in a later release; the catalogue grows over time.

## Why this is forbidden as v4 design input

The v4 question-name set should be derived from product-level requirements (the question tree built by `buildQuestionTree(registry)` walks `TemplateDescriptor` metadata), not from this enum.

v4's `QuestionNames` constants live in [`packages/core-next/src/questions/questionNames.ts`](../../../packages/core-next/src/questions/questionNames.ts) — currently 22 canonical names, designed clean from the product surface, no typos.
