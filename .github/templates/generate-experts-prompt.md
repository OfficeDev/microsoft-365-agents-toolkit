# Expert Generation Prompt

You are an expert knowledge engineer. Your task is to analyze raw source content
and generate a set of **micro-expert markdown files** that teach AI coding agents
domain-specific knowledge.

## Input

You will be given:
1. A directory of raw source files (code, docs, READMEs) synced from an external repo
2. A structural template to follow for each generated expert file

## Instructions

### Phase 1: Analyze the source content

1. Read all files in the synced source directory: `{{SOURCE_PATH}}`
2. Identify the major **domains** (logical groupings of related topics)
3. For each domain, identify **clusters** of related functionality or concepts
4. Note key APIs, patterns, pitfalls, and code examples found in the source

### Phase 2: Plan the expert structure

Create a plan with:
- One **domain directory** per major topic area (e.g., `auth/`, `api/`, `deploy/`)
- One **micro-expert file** per cluster of related functionality
- One **domain index** (`{domain}/index.md`) per domain directory that routes to its experts
- One **root index** (`index.md`) that routes to domain directories

Naming convention: `{domain}/{topic}.md`

### Phase 3: Generate expert files

For each planned expert, create a markdown file at `{{OUTPUT_PATH}}/{domain}/{topic}.md` following the structural template at `{{TEMPLATE_PATH}}`.

Each expert file MUST include:
- **purpose**: One-line description of scope
- **rules**: Numbered list of core rules, patterns, and constraints
- **patterns**: Code examples with correct, idiomatic usage (if applicable)
- **pitfalls**: Common mistakes and how to avoid them (if applicable)

### Phase 4: Generate index files

Create `{{OUTPUT_PATH}}/index.md` as the root router:
```markdown
# expert-router

## purpose
Root task router for {{SKILL_NAME}}. Routes to the correct domain based on the developer's task.

## domains
{{For each domain, add a section with:}}
### {Domain Name}
- **Signals:** {keywords that indicate this domain}
- **Read:** `{domain}/index.md`
```

Create `{{OUTPUT_PATH}}/{domain}/index.md` for each domain:
```markdown
# {domain}-router

## purpose
Routes {domain} tasks to the correct micro-expert.

## task clusters
{{For each cluster:}}
### {Cluster description}
- **When:** {conditions that trigger this cluster}
- **Read:** `{expert-file-1}.md`, `{expert-file-2}.md`

## file inventory
{{Alphabetical list of all expert files in this domain}}
```

### Phase 5: Generate README

Create `{{OUTPUT_PATH}}/README.md` with:
- Expert system overview
- Domain inventory with file counts
- How the routing works
- Scenario-based entry points (common tasks → which experts to load)

## Quality Rules

1. **No hallucinated APIs** — Only reference APIs and patterns found in the source content
2. **Cite source files** — When a rule or pattern comes from a specific source file, reference it
3. **Be specific** — Avoid generic advice. Every rule should be actionable and verifiable
4. **Keep experts focused** — Each file covers ONE cluster of functionality. If it exceeds ~500 lines, split it
5. **Preserve code examples** — If the source has working code, include it verbatim in the expert
6. **Use consistent formatting** — Follow the template structure exactly

## Output

Write all generated files to `{{OUTPUT_PATH}}`. Do not modify any files outside this directory.
After generation, list all created files with a one-line description of each.
