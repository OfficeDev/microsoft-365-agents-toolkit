# Azure SDK Patterns by Language

Reference for creating skills that teach agents to write code following official Azure SDK guidelines.

**Official Documentation:** https://azure.github.io/azure-sdk/

---

## Table of Contents

1. [Core Principles (All Languages)](#core-principles-all-languages)
2. [Standard Naming Conventions](#standard-naming-conventions)
3. [Python Patterns](#python-patterns)
4. [.NET (C#) Patterns](#net-c-patterns)
5. [Java Patterns](#java-patterns)
6. [TypeScript/JavaScript Patterns](#typescriptjavascript-patterns)
7. [Authentication (All Languages)](#authentication-all-languages)
8. [Quick Reference Tables](#quick-reference-tables)

---

## Core Principles (All Languages)

Azure SDKs follow five design principles. Skills should reinforce these:

| Principle | Meaning |
|-----------|---------|
| **Idiomatic** | Follow language conventions; feel natural to developers |
| **Consistent** | APIs feel like a single product from a single team |
| **Approachable** | Great docs, predictable defaults, progressive disclosure |
| **Diagnosable** | Clear logging, errors are actionable and human-readable |
| **Dependable** | No breaking changes without major version bump |

**Consistency Priority:** Language conventions > Service conventions > Cross-language conventions

---

## Standard Naming Conventions

### Namespace/Package Format

`<Azure>.<group>.<service>`

| Group | Area | Examples |
|-------|------|----------|
| `ai` | AI/ML services | `Azure.AI.OpenAI`, `azure-ai-agents` |
| `data` | Databases | `Azure.Data.Cosmos`, `azure-cosmos` |
| `storage` | Storage services | `Azure.Storage.Blobs`, `@azure/storage-blob` |
| `identity` | Auth/Identity | `Azure.Identity`, `azure-identity` |
| `messaging` | Messaging | `Azure.Messaging.ServiceBus` |
| `security` | Security/Crypto | `Azure.Security.KeyVault` |

### Standard Verb Prefixes (All Languages)

| Verb | Behavior | Returns |
|------|----------|---------|
| `create` | Create new; fail if exists | Created item |
| `upsert` | Create or update (database-like) | Item |
| `set` | Create or update (dictionary-like) | Item |
| `update` | Fail if doesn't exist | Updated item |
| `get` | Retrieve single; error if missing | Item |
| `list` | Return collection (empty if none) | Pageable |
| `delete` | Succeed even if doesn't exist | void/None |
| `exists` | Check existence | boolean |
| `begin` | Start long-running operation | Poller |

---

## Python Patterns

### Client Naming

```python
# Sync client
class ConfigurationClient:
    pass

# Async client - use Async prefix
class AsyncConfigurationClient:
    pass
```

### Pagination: ItemPaged / AsyncItemPaged

```python
from azure.core.paging import ItemPaged

# Sync iteration
for item in client.list_items():
    print(item.name)

# Page-by-page
for page in client.list_items().by_page():
    for item in page:
        print(item.name)

# With continuation token
for page in client.list_items().by_page(continuation_token="..."):
    print(page)

# Async iteration
async for item in async_client.list_items():
    print(item.name)
```

### Long-Running Operations: LROPoller / AsyncLROPoller

```python
from azure.core.polling import LROPoller

# Start LRO
poller: LROPoller[Result] = client.begin_create_resource(config)

# Check status
if poller.done():
    result = poller.result()

# Wait with timeout
result = poller.result(timeout=60)

# Async LRO
async_poller = await async_client.begin_create_resource(config)
result = await async_poller.result()
```

### Context Managers

```python
# Recommended pattern
with ConfigurationClient(endpoint, credential) as client:
    setting = client.get_setting("key")

# Async
async with AsyncConfigurationClient(endpoint, credential) as client:
    setting = await client.get_setting("key")
```

### Error Handling

```python
from azure.core.exceptions import (
    ResourceNotFoundError,
    ResourceExistsError,
    HttpResponseError,
)

try:
    item = client.get_item("key")
except ResourceNotFoundError:
    print("Not found")
except HttpResponseError as e:
    print(f"HTTP {e.status_code}: {e.message}")
```

### Docstring Format (Sphinx-style)

```python
def get_setting(self, key: str, **kwargs) -> "ConfigurationSetting":
    """Retrieve a configuration setting.

    :param key: The key of the setting.
    :type key: str
    :keyword timeout: Operation timeout in seconds.
    :paramtype timeout: int
    :returns: The configuration setting.
    :rtype: ~azure.appconfig.ConfigurationSetting
    :raises ~azure.core.exceptions.ResourceNotFoundError: If setting not found.
    """
```

---

## .NET (C#) Patterns

### Client Naming

```csharp
namespace Azure.Data.Configuration
{
    // Service client with Client suffix
    public class ConfigurationClient { }
    
    // Options class
    public class ConfigurationClientOptions : ClientOptions { }
}
```

### Response Wrapper: Response<T>

```csharp
// Single item
public Response<ConfigurationSetting> GetSetting(string key);
public Task<Response<ConfigurationSetting>> GetSettingAsync(string key);

// No content
public Response DeleteSetting(string key);
public Task<Response> DeleteSettingAsync(string key);
```

### Pagination: Pageable<T> / AsyncPageable<T>

```csharp
// Sync
foreach (ConfigurationSetting setting in client.GetSettings())
{
    Console.WriteLine(setting.Key);
}

// Async
await foreach (ConfigurationSetting setting in client.GetSettingsAsync())
{
    Console.WriteLine(setting.Key);
}
```

### Long-Running Operations: Operation<T>

```csharp
// With WaitUntil parameter
Operation<AnalyzeResult> operation = await client.StartAnalyzeAsync(
    WaitUntil.Completed,  // or WaitUntil.Started
    document);

AnalyzeResult result = operation.Value;

// Manual polling
Operation<AnalyzeResult> operation = await client.StartAnalyzeAsync(
    WaitUntil.Started, document);

while (!operation.HasCompleted)
{
    await operation.UpdateStatusAsync();
    await Task.Delay(1000);
}
```

### Mocking Support

```csharp
public class ConfigurationClient
{
    // Protected parameterless constructor for mocking
    protected ConfigurationClient() { }
    
    // Virtual methods for mocking
    public virtual Response<ConfigurationSetting> GetSetting(string key);
}
```

### Error Handling

```csharp
try
{
    var setting = await client.GetSettingAsync("key");
}
catch (RequestFailedException ex) when (ex.Status == 404)
{
    Console.WriteLine("Not found");
}
catch (RequestFailedException ex)
{
    Console.WriteLine($"Error: {ex.Status} - {ex.ErrorCode}");
}
```

---

## Java Patterns

### Client Naming

```java
// Sync client
public final class ConfigurationClient { }

// Async client
public final class ConfigurationAsyncClient { }

// Builder (the ONLY way to create clients)
public final class ConfigurationClientBuilder {
    public ConfigurationClient buildClient() { }
    public ConfigurationAsyncClient buildAsyncClient() { }
}
```

### Builder Pattern

```java
ConfigurationClient client = new ConfigurationClientBuilder()
    .endpoint(endpoint)
    .credential(new DefaultAzureCredentialBuilder().build())
    .serviceVersion(ConfigurationServiceVersion.V2023_10_01)
    .buildClient();
```

### Pagination: PagedIterable<T> / PagedFlux<T>

```java
// Sync - standard for loop
for (ConfigurationSetting setting : client.listSettings()) {
    System.out.println(setting.getKey());
}

// Sync - Stream API
client.listSettings().stream()
    .filter(s -> s.getKey().startsWith("app"))
    .forEach(System.out::println);

// Async - Reactor
client.listSettings()
    .subscribe(setting -> System.out.println(setting.getKey()));
```

### Long-Running Operations: SyncPoller<T,U> / PollerFlux<T,U>

```java
// Sync
SyncPoller<OperationResult, AnalyzeResult> poller = 
    client.beginAnalyze(document);
poller.waitForCompletion();
AnalyzeResult result = poller.getFinalResult();

// Async
client.beginAnalyze(document)
    .last()
    .flatMap(AsyncPollResponse::getFinalResult)
    .subscribe(result -> System.out.println(result));
```

### Reactor Types

| Type | Purpose |
|------|---------|
| `Mono<T>` | 0 or 1 item |
| `Flux<T>` | 0 to N items |
| `PagedFlux<T>` | Paginated collections |
| `PollerFlux<T,U>` | Long-running operations |

### Annotations

```java
@ServiceClient(builder = ConfigurationClientBuilder.class)
public final class ConfigurationClient {
    
    @ServiceMethod(returns = ReturnType.SINGLE)
    public ConfigurationSetting getSetting(String key) { }
    
    @ServiceMethod(returns = ReturnType.COLLECTION)
    public PagedIterable<ConfigurationSetting> listSettings() { }
}
```

---

## TypeScript/JavaScript Patterns

### Package Naming

```typescript
// Package: @azure/service-name (kebab-case)
// Client: ServiceClient (PascalCase with Client suffix)

import { ServiceClient } from "@azure/service-name";
```

### Pagination: PagedAsyncIterableIterator

```typescript
// Iterate items
for await (const item of client.listItems()) {
    console.log(item.name);
}

// Iterate by page
for await (const page of client.listItems().byPage()) {
    console.log(`Page has ${page.length} items`);
}

// With continuation token
const iterator = client.listItems().byPage({ continuationToken });
```

### Long-Running Operations

```typescript
// Methods starting LRO use 'begin' prefix
const poller = await client.beginAnalyzeDocument(modelId, document, {
    pollInterval: 2000
});

// Wait for completion
const result = await poller.pollUntilDone();

// Serialize state for later
const state = poller.toString();
const restored = await client.beginAnalyzeDocument(modelId, document, {
    resumeFrom: state
});
```

### Cancellation: AbortSignal

```typescript
import { AbortController } from "@azure/abort-controller";

const controller = new AbortController();
setTimeout(() => controller.abort(), 5000);

try {
    const item = await client.createItem({
        abortSignal: controller.signal
    });
} catch (e) {
    if (e.name === "AbortError") {
        console.log("Cancelled");
    }
}
```

### Options Pattern

```typescript
interface CreateItemOptions {
    abortSignal?: AbortSignalLike;
    timeoutInMs?: number;        // Duration suffix: InMs, InSeconds
    onlyIfChanged?: boolean;     // Conditional request
}
```

### Error Handling

```typescript
import { RestError } from "@azure/core-rest-pipeline";

try {
    await client.createItem(item);
} catch (e) {
    // Check name, not instanceof
    if (e.name === "RestError") {
        console.error(`HTTP ${e.statusCode}: ${e.message}`);
    }
}
```

---

## Authentication (All Languages)

**Always use `DefaultAzureCredential` as the primary pattern:**

### Python

```python
from azure.identity import DefaultAzureCredential
credential = DefaultAzureCredential()
client = ServiceClient(endpoint, credential)
```

### .NET

```csharp
var credential = new DefaultAzureCredential();
var client = new ServiceClient(new Uri(endpoint), credential);
```

### Java

```java
TokenCredential credential = new DefaultAzureCredentialBuilder().build();
ServiceClient client = new ServiceClientBuilder()
    .endpoint(endpoint)
    .credential(credential)
    .buildClient();
```

### TypeScript

```typescript
import { DefaultAzureCredential } from "@azure/identity";
const credential = new DefaultAzureCredential();
const client = new ServiceClient(endpoint, credential);
```

**Rules:**
- Never hardcode credentials
- Never persist/cache tokens manually (credential handles refresh)
- Use environment variables for configuration

---

## Quick Reference Tables

### Client Types by Language

| Pattern | Python | .NET | Java | TypeScript |
|---------|--------|------|------|------------|
| Sync Client | `Client` | `Client` | `Client` | `Client` |
| Async Client | `AsyncClient` | N/A (Async methods) | `AsyncClient` | N/A (Promise) |
| Builder | N/A | N/A | `ClientBuilder` | N/A |

### Pagination Types

| Language | Sync | Async |
|----------|------|-------|
| Python | `ItemPaged[T]` | `AsyncItemPaged[T]` |
| .NET | `Pageable<T>` | `AsyncPageable<T>` |
| Java | `PagedIterable<T>` | `PagedFlux<T>` |
| TypeScript | N/A | `PagedAsyncIterableIterator<T>` |

### LRO Types

| Language | Sync | Async |
|----------|------|-------|
| Python | `LROPoller[T]` | `AsyncLROPoller[T]` |
| .NET | `Operation<T>` | `Operation<T>` |
| Java | `SyncPoller<T,U>` | `PollerFlux<T,U>` |
| TypeScript | N/A | `PollerLike<T>` |

### Response Wrappers

| Language | Single Item | Collection |
|----------|-------------|------------|
| Python | Direct return | `ItemPaged[T]` |
| .NET | `Response<T>` | `Pageable<T>` |
| Java | Direct return | `PagedIterable<T>` |
| TypeScript | `Promise<T>` | `PagedAsyncIterableIterator<T>` |

---

## Official Documentation Links

- **General Guidelines:** https://azure.github.io/azure-sdk/general_introduction.html
- **Python:** https://azure.github.io/azure-sdk/python_design.html
- **.NET:** https://azure.github.io/azure-sdk/dotnet_introduction.html
- **Java:** https://azure.github.io/azure-sdk/java_introduction.html
- **TypeScript:** https://azure.github.io/azure-sdk/typescript_introduction.html

When creating Azure SDK skills, reference these docs via the `microsoft-docs` MCP for current API signatures.
