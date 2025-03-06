{
    {{^EmbeddedKnowledgeEnabled}}
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.2/schema.json",
    "version": "v1.2",
    {{/EmbeddedKnowledgeEnabled}}
    {{#EmbeddedKnowledgeEnabled}}
    "version": "v1.3",
    {{/EmbeddedKnowledgeEnabled}}
    "name": "{{appName}}",
    "description": "Declarative agent created with Teams Toolkit can assist user in calling APIs and retrieving responses",
    "instructions": "$[file('instruction.txt')]"
}
