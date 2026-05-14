# Output Patterns

Patterns for producing consistent, high-quality output in skills.

## Template Pattern

Provide templates for output format. Match strictness to requirements.

**For strict requirements (API responses, data formats):**

```markdown
## Report structure

ALWAYS use this exact template structure:

# [Analysis Title]

## Executive summary
[One-paragraph overview of key findings]

## Key findings
- Finding 1 with supporting data
- Finding 2 with supporting data
- Finding 3 with supporting data

## Recommendations
1. Specific actionable recommendation
2. Specific actionable recommendation
```

**For flexible guidance:**

```markdown
## Report structure

Sensible default format; adapt as needed:

# [Analysis Title]

## Executive summary
[Overview]

## Key findings
[Adapt sections based on what you discover]

## Recommendations
[Tailor to the specific context]
```

## Examples Pattern

For output quality dependent on examples, provide input/output pairs:

```markdown
## Commit message format

Generate commit messages following these examples:

**Example 1:**
Input: Added user authentication with JWT tokens
Output:
feat(auth): implement JWT-based authentication

Add login endpoint and token validation middleware

**Example 2:**
Input: Fixed bug where dates displayed incorrectly in reports
Output:
fix(reports): correct date formatting in timezone conversion

Use UTC timestamps consistently across report generation

Follow this style: type(scope): brief description, then detailed explanation.
```

Examples help agents understand desired style more clearly than descriptions alone.

## Azure SDK Code Patterns

### Client Initialization Template

```python
# Standard Azure SDK client setup
import os
from azure.identity import DefaultAzureCredential
from azure.<service> import <Service>Client

credential = DefaultAzureCredential()
client = <Service>Client(
    endpoint=os.environ["AZURE_<SERVICE>_ENDPOINT"],
    credential=credential
)
```

### CRUD Method Template

```python
# Create
item = client.create_<noun>(
    name="example",
    config=<Noun>Config(
        property1="value1",
        property2="value2"
    )
)

# Read
item = client.get_<noun>(item_id)

# List (with pagination)
for item in client.list_<nouns>():
    print(item.name)

# Update
updated = client.update_<noun>(item_id, new_config)

# Delete
client.delete_<noun>(item_id)
```

### Async Client Template

```python
import asyncio
from azure.identity.aio import DefaultAzureCredential
from azure.<service>.aio import <Service>Client

async def main():
    credential = DefaultAzureCredential()
    async with <Service>Client(endpoint, credential) as client:
        # Async operations
        item = await client.get_<noun>(item_id)
        
        # Async pagination
        async for item in client.list_<nouns>():
            print(item.name)

asyncio.run(main())
```

### Error Handling Template

```python
from azure.core.exceptions import (
    ResourceNotFoundError,
    ResourceExistsError,
    HttpResponseError,
)

try:
    result = client.get_<noun>(item_id)
except ResourceNotFoundError:
    # Handle 404
    print(f"Resource {item_id} not found")
except ResourceExistsError:
    # Handle 409
    print(f"Resource already exists")
except HttpResponseError as e:
    # Handle other HTTP errors
    print(f"HTTP {e.status_code}: {e.message}")
```

### Feature Comparison Table Template

```markdown
## Clients

| Client | Purpose | When to Use |
|--------|---------|-------------|
| `ServiceClient` | Core operations | Standard use cases |
| `AsyncServiceClient` | Async operations | High-throughput scenarios |
| `ServiceAdminClient` | Management | Creating/deleting resources |
```

### Environment Variables Template

```markdown
## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_<SERVICE>_ENDPOINT` | Yes | Service endpoint URL |
| `AZURE_<SERVICE>_KEY` | No | API key (alternative to DefaultAzureCredential) |
| `AZURE_CLIENT_ID` | No | For service principal auth |
| `AZURE_TENANT_ID` | No | For service principal auth |
| `AZURE_CLIENT_SECRET` | No | For service principal auth |
```
