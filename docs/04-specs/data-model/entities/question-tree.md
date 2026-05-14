# Question tree

A tree of `UserInputQuestion` nodes consumed by all surfaces (VS Code, VS, CLI). See [01-product/ux/question-model.md](../../01-product/ux/question-model.md) for the UX angle; this page is the data structure.

## v4 types

Source: [`packages/core-next/src/api/qm/`](../../../packages/core-next/src/api/qm/) (merged from v3's `@microsoft/teamsfx-api`).

```typescript
type IQTreeNode = {
  data: UserInputQuestion;
  children?: Array<{
    condition: Condition;     // when to activate this child
    node: IQTreeNode;
  }>;
};

type UserInputQuestion =
  | TextInputQuestion
  | SingleSelectQuestion
  | MultiSelectQuestion
  | SingleFileQuestion
  | MultiFileQuestion
  | FolderQuestion
  | ConfirmQuestion;

type Condition =
  | { equals: any }
  | { enum: any[] }
  | { contains: any }
  | ConditionFunc;
```

## Common factories

`packages/core-next/src/questions/commonQuestions.ts` (18 factories):

| Factory | Returns |
|---------|---------|
| `projectNameQuestion()` | TextInputQuestion with name validation |
| `languageQuestion(allowed)` | SingleSelectQuestion of "typescript" / "javascript" / ... |
| `folderQuestion()` | FolderQuestion |
| `appNameQuestion()` | TextInputQuestion |
| `llmProviderQuestion()` | SingleSelectQuestion (OpenAI / AzureOpenAI / Foundry / ...) |
| `foundryEndpointQuestion()` | TextInputQuestion |
| `foundryAgentIdQuestion()` | TextInputQuestion |
| `graphConnectorTenantIdQuestion()` | TextInputQuestion |
| ... | ... |

These compose with template-specific questions declared in `TemplateDescriptor.questions`.

## Building a tree

`buildQuestionTree(registry)` walks the `TemplateRegistry` and produces:

```
projectType
└── (if projectType=agent) templateCategory
    ├── (if category=da) templateId
    │   ├── (if templateId=da/basic) appName → folder
    │   └── (if templateId=da/api-plugin-oauth) appName → folder
    ├── (if category=cea) templateId
    │   └── (if templateId=cea/weather) language → appName → folder
    ├── (if category=ai) templateId
    │   └── (if templateId=ai/chat-bot) llmProvider → language → appName → folder
    └── ...
```

Each leaf-template's `questions?: QuestionSpec[]` are spliced into the tree at the right place.

## Traversal

`traverseQuestionTree(tree, ui, inputs)` is an iterative DFS with:

- **Back-stack** — user can answer "back" to revisit the previous question.
- **Pre-filled bypass** — if `inputs[questionName]` is already set (e.g. from CLI flag), the question is skipped.
- **Subtree-aware skipping** — when a parent answer eliminates a subtree, the entire subtree is skipped, not just the immediate next question.
- **Condition evaluation** — `equals`, `enum`, `contains`, or any `ConditionFunc`.

Result: `inputs` populated with every answered question. Then `createProjectOp(ctx, inputs)` does the scaffolding.

## CLI mapping

Every question's `cliName`, `cliShortName`, `cliDescription`, `isBoolean` fields drive `mapQuestionToOption()` so the same tree drives `--<flag>` parsing under Commander.

A question with no CLI metadata is a runtime warning — it cannot be answered headlessly.
