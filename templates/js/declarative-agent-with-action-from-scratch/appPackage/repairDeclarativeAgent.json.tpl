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
    "name": "{{appName}}${{APP_NAME_SUFFIX}}",
    "description": "This declarative agent helps you with finding car repair records.",
    "instructions": "$[file('instruction.txt')]",
    "conversation_starters": [
        {
            "text": "Show repair records assigned to Karin Blair"
        }
    ],
    "actions": [
        {
            "id": "repairPlugin",
            "file": "ai-plugin.json"
        }
    ]
}
