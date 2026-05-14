# Question model engine

Source: [`packages/core-next/src/questions/`](../../../packages/core-next/src/questions/).

## Layout

```
questions/
├── questionNames.ts    — QuestionNames constants (22)
├── commonQuestions.ts  — 18 reusable factories
├── treeBuilder.ts      — buildQuestionTree(registry)
├── traverse.ts         — traverseQuestionTree(tree, ui, inputs)
└── index.ts            — barrel
```

## Tree construction

`buildQuestionTree(registry)` walks the `TemplateRegistry`:

```
projectType (root)
└── (if projectType=agent) templateCategory
    ├── (if category=da)  templateId
    │   ├── (if templateId=da/basic)            appName → folder
    │   ├── (if templateId=da/api-plugin-oauth) appName → folder
    │   └── ...
    ├── (if category=cea) templateId
    │   ├── (if templateId=cea/basic)   language → appName → folder
    │   └── (if templateId=cea/weather) language → appName → folder
    ├── (if category=ai)  templateId
    │   └── (if templateId=ai/chat-bot) llmProvider → language → appName → folder
    └── ...
```

Per-template extra questions declared in `TemplateDescriptor.questions` are spliced into the tree at the right place automatically.

## Traversal

`traverseQuestionTree(tree, ui, inputs)` is iterative DFS with:

| Feature | Effect |
|---------|--------|
| Back-stack | User can return to previous question |
| Pre-filled bypass | If `inputs[questionName]` is already set (CLI flag), skip |
| Subtree-aware skip | When parent answer eliminates a subtree, skip the entire subtree |
| Condition evaluation | `equals`, `enum`, `contains`, or `ConditionFunc` |

Result: `inputs` populated with every answered question. `createProjectInteractive(ctx, inputs)` then calls `createProjectOp`.

## Common factories

`commonQuestions.ts` (18):

| Factory | Returns |
|---------|---------|
| `projectNameQuestion()` | `TextInputQuestion` |
| `appNameQuestion()` | `TextInputQuestion` |
| `folderQuestion()` | `FolderQuestion` |
| `languageQuestion(allowed)` | `SingleSelectQuestion` |
| `llmProviderQuestion()` | `SingleSelectQuestion` (OpenAI / AzureOpenAI / Foundry / ...) |
| `foundryEndpointQuestion()` | `TextInputQuestion` |
| `foundryAgentIdQuestion()` | `TextInputQuestion` |
| `graphConnectorTenantIdQuestion()` | `TextInputQuestion` |
| `confirmQuestion(message)` | `ConfirmQuestion` |
| ... | ... |

## CLI mapping

Every question's `cliName`, `cliShortName`, `cliDescription`, `isBoolean` fields drive `mapQuestionToOption()` so the same tree drives `--<flag>` parsing under Commander. A question with no CLI metadata is a runtime warning — it cannot be answered headlessly.

## Tests

- `tests/unit/questions/commonQuestions.test.ts` — 15 factory tests.
- `tests/unit/questions/treeBuilder.test.ts` — registry → tree.
- `tests/unit/questions/traverse.test.ts` — 20 unit tests.
- `tests/unit/questions/traverseIntegration.test.ts` — build + traverse integration (3 tests).
