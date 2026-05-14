# Visual Studio surface

The Visual Studio extension lives outside this repository, but consumes packages built here:

| Package | Used for |
|---------|----------|
| [`packages/dotnet-sdk`](../../../packages/dotnet-sdk/) | C# SDK for SSO and Teams app helpers |
| [`packages/function-extension`](../../../packages/function-extension/) | Authentication-aware Azure Function bindings |
| [`packages/server`](../../05-engineering/package-reference/server.md) | JSON-RPC server bridging Visual Studio to fx-core |

## Communication model

VS hosts the extension UI; long-running engine work runs in a Node child process driven by `packages/server` (JSON-RPC over stdio). VS sends method calls (`createProject`, `provisionResources`, `deploy`, `publish`) and receives progress notifications and error responses on the same channel.

## Project templates

VS dialogs surface the templates registered in fx-core's generator catalogue, filtered to those marked for VS (`platform: VS` in metadata). Languages: C# only.

## Auth

VS uses its own MSAL flow when possible (`Microsoft.VisualStudio.Services.Client` MSAL integrations), falling back to the same broker plugin paths the CLI uses on Windows.

## Error display

The same `FxError` shape flows back to VS via JSON-RPC; VS surfaces `displayMessage` in dialogs and the Tasks list, with `helpLink` becoming a clickable affordance.
