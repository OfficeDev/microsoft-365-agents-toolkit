# Question model

The toolkit needs the same answer to "what do you want to build?" regardless of whether the user is clicking a webview, navigating a CLI prompt, or filling a Visual Studio dialog. The **question model** is the abstraction that makes this possible.

## Shape

A *question* is a node:

```typescript
type UserInputQuestion =
  | TextInputQuestion
  | SingleSelectQuestion
  | MultiSelectQuestion
  | SingleFileQuestion
  | MultiFileQuestion
  | FolderQuestion
  | ConfirmQuestion;
```

A *tree* is a node with children whose visibility is conditioned on parent answers.

## Surfaces consume the same tree

| Surface | Driver |
|---------|--------|
| VS Code | `qm/` module renders Quick Pick / input box / open dialog |
| Visual Studio | Native WPF dialogs |
| CLI v3 | `CLIUserInteraction` with `@inquirer/prompts` |
| CLI v4 | `CLIUserInteraction` (in cli-next) |

All consume the same tree definitions exported by the engine (`fx-core/src/question/scaffold/` v3, `core-next/src/questions/` v4).

## v4 simplifications

`packages/core-next/src/questions/` provides:

- `QuestionNames` — 22 canonical constants (no magic strings).
- `commonQuestions.ts` — 18 reusable factories (`projectNameQuestion()`, `languageQuestion()`, `llmProviderQuestion()`, ...).
- `buildQuestionTree(registry)` — auto-generates an `IQTreeNode` from `TemplateRegistry` metadata.
- `traverseQuestionTree(tree, ui, inputs)` — iterative DFS with a back-stack and pre-filled-input bypass.
- `createProjectInteractive(ctx, inputs)` — combines the two for one-shot interactive create.

This means a new template auto-gets:

- A CLI subcommand (via `buildNewCommands`).
- A question tree node (via `buildQuestionTree`).
- A question traversal that lets the user back out of subtrees.

See [`questions.test.ts`](../../packages/core-next/tests/unit/questions/) for the test surface.

## Conditions

A child node activates only when the parent answer matches a condition:

| Condition kind | Example |
|----------------|---------|
| `equals` | Activate "Python framework" only if `language === "python"` |
| `enum` | Activate "AI provider" only if template ∈ `[ai-chat, weather, rag]` |
| `contains` | List membership |
| `ConditionFunc` | Arbitrary predicate (used sparingly) |

## CLI mapping

In CLI mode, every `Question` carries CLI metadata used by `mapQuestionToOption()`:

| Field | CLI usage |
|-------|-----------|
| `cliName` | `--<name>` long flag |
| `cliShortName` | `-<x>` short flag |
| `cliDescription` | Help text |
| `isBoolean` | `--<flag>` no value |

Adding a question that doesn't set CLI metadata produces a runtime warning — every question must work in both interactive and non-interactive modes.

## Validation

Per question:

- `validation` — declarative (`StringValidation`: pattern, enum, minLength).
- `validationFunc` — programmatic (`ValidateFunc<T>`).
- `default` / `defaultDynamic` — pre-filled value.

Validation failures re-prompt in interactive mode; in CI they surface as `InvalidChoiceError` or `MissingRequiredOptionError`.
