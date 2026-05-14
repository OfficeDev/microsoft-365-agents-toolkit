# `packages/server` — `@microsoft/teamsfx-server`

JSON-RPC bridge between the **Visual Studio extension** and `fx-core`.

## Role

VS hosts the extension UI in WPF / .NET. Long-running engine work runs in a Node child process driven by this package over JSON-RPC stdio.

## Communication model

- VS sends method calls: `createProject`, `provisionResources`, `deploy`, `publish`, `getQuestions`, `getProjectInfo`, ...
- Server delegates to `fx-core` and streams back results.
- Progress and log events flow back as JSON-RPC notifications.
- Errors flow back as `FxError` shapes (`source`, `name`, `message`, `displayMessage`, `helpLink`).

## Bundler

Webpack (legacy). No esbuild migration planned until VS Code extension migration completes.

## Status

Stable. Not part of the v4 migration scope. When v4 reaches the VS extension, this server will be re-pointed at `core-next` operations.
