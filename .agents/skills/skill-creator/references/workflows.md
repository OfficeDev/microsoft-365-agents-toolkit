# Workflow Patterns

Patterns for structuring multi-step processes in skills.

## Sequential Workflows

For complex tasks, break operations into clear steps. Provide an overview at the start:

```markdown
Filling a PDF form involves these steps:

1. Analyze the form (run analyze_form.py)
2. Create field mapping (edit fields.json)
3. Validate mapping (run validate_fields.py)
4. Fill the form (run fill_form.py)
5. Verify output (run verify_output.py)
```

## Conditional Workflows

For tasks with branching logic, guide through decision points:

```markdown
1. Determine the modification type:
   **Creating new content?** → Follow "Creation workflow" below
   **Editing existing content?** → Follow "Editing workflow" below

2. Creation workflow: [steps]
3. Editing workflow: [steps]
```

## Azure SDK Workflows

### CRUD Lifecycle Pattern

```markdown
## Working with [Resource]

### Create
\`\`\`python
resource = client.create_resource(name="example", config={...})
\`\`\`

### Read
\`\`\`python
# Single item
resource = client.get_resource("resource-id")

# List with pagination
for resource in client.list_resources():
    print(resource.name)
\`\`\`

### Update
\`\`\`python
resource = client.update_resource("resource-id", new_config={...})
\`\`\`

### Delete
\`\`\`python
client.delete_resource("resource-id")
\`\`\`
```

### Long-Running Operation Pattern

```markdown
## Processing [Resource]

Long-running operations use the poller pattern:

\`\`\`python
# Start operation
poller = client.begin_process_resource(resource_id, config)

# Option 1: Wait for completion
result = poller.result()

# Option 2: Poll with status updates
while not poller.done():
    print(f"Status: {poller.status()}")
    time.sleep(5)
result = poller.result()

# Option 3: Use callback
poller.add_done_callback(lambda r: print(f"Done: {r}"))
\`\`\`
```

### Agent Lifecycle Pattern (Azure AI Agents)

```markdown
## Agent Workflow

1. **Create Agent** with tools and instructions
2. **Create Thread** for conversation
3. **Add Messages** to thread
4. **Run Agent** on thread
5. **Process Response** (handle tool calls if needed)
6. **Cleanup** - delete agent when done

\`\`\`python
# 1. Create
agent = client.create_agent(model="gpt-4o", instructions="...")

# 2-4. Thread, Message, Run
thread = client.threads.create()
client.messages.create(thread_id=thread.id, content="...")
run = client.runs.create(thread_id=thread.id, agent_id=agent.id)

# 5. Wait for completion
while run.status in ["queued", "in_progress"]:
    run = client.runs.retrieve(thread_id=thread.id, run_id=run.id)
    time.sleep(1)

# 6. Cleanup
client.delete_agent(agent.id)
\`\`\`
```

### Error Recovery Pattern

```markdown
## Error Handling

\`\`\`python
from azure.core.exceptions import (
    ResourceNotFoundError,
    ResourceExistsError,
    HttpResponseError,
)

try:
    result = client.get_resource("id")
except ResourceNotFoundError:
    # Handle missing resource
    result = client.create_resource("id", default_config)
except HttpResponseError as e:
    if e.status_code == 429:  # Rate limited
        time.sleep(e.retry_after or 60)
        result = client.get_resource("id")
    else:
        raise
\`\`\`
```
