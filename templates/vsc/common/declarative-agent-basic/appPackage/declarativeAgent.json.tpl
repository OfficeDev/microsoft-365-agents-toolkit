{
    {{^EmbeddedKnowledgeEnabled}}
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.4/schema.json",
    "version": "v1.4",
    {{/EmbeddedKnowledgeEnabled}}
    {{#EmbeddedKnowledgeEnabled}}
    "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.6/schema.json",
    "version": "v1.6",
    {{/EmbeddedKnowledgeEnabled}}
    {{#SensitivityLabelEnabled}}
    "sensitivity_label": {
        "id": ""
    },
    {{/SensitivityLabelEnabled}}
    "name": "{{appName}}",
    "description": "Declarative agent created with Microsoft 365 Agents Toolkit",
    {{^CopilotConnector}}
    "instructions": "$[file('instruction.txt')]"
    {{/CopilotConnector}}
    {{#CopilotConnector}}
    "instructions": "$[file('instruction.txt')]",
    "capabilities": [
        {
            "name": "GraphConnectors",
            "connections": [
                {
                    "connection_id": "${{CONNECTOR_ID}}"
                }
            ]
        }
    ]
    {{/CopilotConnector}}
}