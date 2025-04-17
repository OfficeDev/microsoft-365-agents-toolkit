{
    {{^EmbeddedKnowledgeEnabled}}
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.3/schema.json",
    "version": "v1.3",
    {{/EmbeddedKnowledgeEnabled}}
    {{#EmbeddedKnowledgeEnabled}}
    "version": "v1.4",
    {{/EmbeddedKnowledgeEnabled}}
    {{#SensitivityLabelEnabled}}
    "sensitivity_label": "",
    {{/SensitivityLabelEnabled}}
    "name": "{{appName}}",
    "description": "Declarative agent created with Microsoft 365 Agents Toolkit can assist user in calling APIs and retrieving responses",
    "instructions": "$[file('instruction.txt')]"
}
