# 6 — Runtime views

Selected sequences for the most-trafficked operations. v4 (`core-next`) shown — v3 sequences are equivalent in shape.

## Create project (interactive)

```mermaid
sequenceDiagram
    participant User
    participant CLI as cli-next
    participant Q as questions/traverse
    participant TR as TemplateRegistry
    participant SC as scaffold/scaffolder
    participant FS as Filesystem

    User->>CLI: atk new
    CLI->>CLI: createCliContext()
    CLI->>Q: buildQuestionTree(registry)
    CLI->>Q: traverseQuestionTree(tree, ui, inputs)
    Q-->>User: prompt category → template → language → extras
    User-->>Q: answers
    CLI->>TR: registry.lookup(templateId)
    TR-->>CLI: TemplateDescriptor
    CLI->>SC: scaffoldTemplates(descriptor, replaceMap)
    SC->>SC: download.fetchZip OR loadLocalFallback
    SC->>SC: unzipWithTransform (filter prefix, render Mustache)
    SC->>FS: write files
    CLI-->>User: PostAction[]
```

## Provision (v4 `provisionOp`)

```mermaid
sequenceDiagram
    participant User
    participant CLI as cli-next provisionAction
    participant Op as provisionOp
    participant Pre as prerequisites
    participant Yaml as parser/resolver
    participant Exec as executor
    participant Drv as drivers
    participant Env as environment

    User->>CLI: atk provision --env dev
    CLI->>Op: runOperation(provisionOp, {projectPath, envName: "dev"})
    Op->>Env: loadEnv("dev")
    Op->>Yaml: parseProjectYaml + resolveLifecycle
    Op->>Op: analyzeSteps (needsM365? needsAzure?)
    Op->>Pre: ensureM365Auth → ensureAzureAuth
    Pre-->>Op: tokens, tenantId
    Op->>Pre: ensureSubscription → ensureResourceGroup → ensureResourceSuffix
    Op->>Pre: confirmProvision (dialog)
    Op->>Exec: executeLifecycle(steps, envMap)
    loop each step
        Exec->>Drv: validate(input) → execute(input)
        Drv-->>Exec: outputs
        Exec->>Exec: merge outputs into envMap
    end
    Exec-->>Op: result
    Op->>Env: persistEnv("dev", envMap)
    Op-->>CLI: PostAction[] (open Azure portal, etc.)
```

Source: [`packages/core-next/src/lifecycle/operations.ts`](../../packages/core-next/src/lifecycle/operations.ts)

## Deploy

Same shape as provision but with `ensureM365Auth + confirmDeploy` (no Azure auth gate; deploy actions handle auth). For local/testtool/playground/sandbox env names, `confirmDeploy` is skipped.

```mermaid
sequenceDiagram
    participant CLI as deployAction
    participant Op as deployOp
    participant Exec as executor
    participant Build as cli/runNpmCommand
    participant Push as azureFunctions/zipDeploy

    CLI->>Op: runOperation(deployOp, {projectPath, envName})
    Op->>Op: confirmDeploy (skipped for local-ish envs)
    Op->>Exec: executeLifecycle(deploy steps)
    Exec->>Build: npm run build
    Build-->>Exec: artifacts
    Exec->>Push: zip(src) → POST /api/zipdeploy
    Push-->>Exec: deployment URL
    Exec-->>Op: result
```

## Publish

```mermaid
sequenceDiagram
    participant CLI as publishAction
    participant Op as publishOp
    participant Zip as teamsApp/zipAppPackage
    participant Pub as teamsApp/publishAppPackage
    participant Graph as MS Graph

    CLI->>Op: runOperation(publishOp, {projectPath, envName})
    Op->>Zip: bundle manifest + icons
    Op->>Pub: POST /beta/appCatalogs/teamsApps (or PUT to update)
    Pub->>Graph: HTTP
    Graph-->>Pub: 201/200
    Pub-->>Op: appId, status
    Op-->>CLI: PostAction (admin URL)
```

## Driver execution detail

Inside `executeLifecycle`, before each driver call:

1. `ctx.projectPath` is injected as `PROJECT_PATH` if absent.
2. envMap entries are temporarily synced into `process.env` so drivers loading external files (ARM parameters, AAD manifest templates) can resolve `${{VAR}}` placeholders produced by earlier steps.
3. The driver runs.
4. `finally` block removes the injected env vars to avoid state leaking between steps.

Source: [`packages/core-next/src/lifecycle/executor.ts`](../../packages/core-next/src/lifecycle/executor.ts)
