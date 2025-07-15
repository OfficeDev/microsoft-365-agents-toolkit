{
  "$schema": "https://aka.ms/json-schemas/copilot-extensions/v2.1/plugin.schema.json",
  "schema_version": "v2.4",
  "name_for_human": "{{appName}}",
  "description_for_human": "{{appName}}${{APP_NAME_SUFFIX}}",
  "contact_email": "publisher-email@example.com",
  "namespace": "{{appName}}",
  "functions": [
        {
            "name": "microsoft_docs_search",
            "description": "Search official Microsoft/Azure documentation to find the most relevant and trustworthy content for a user's query. This tool returns up to 10 high-quality content chunks (each max 500 tokens), extracted from Microsoft Learn and other official sources. Each result includes the article title, URL, and a self-contained content excerpt optimized for fast retrieval and reasoning. Always use this tool to quickly ground your answers in accurate, first-party Microsoft/Azure knowledge.",
            "parameters": {
                "type": "object",
                "properties": {
                    "question": {
                        "description": "a question or topic about Microsoft/Azure products, services, platforms, developer tools, frameworks, or APIs",
                        "type": "string"
                    }
                },
                "required": [
                    "question"
                ]
            }
        }
    ],
    "runtimes": [
        {
            "type": "RemoteMCPServer",
            "spec": {
                "url": "https://learn.microsoft.com/api/mcp",
                "enable_dynamic_discovery": false
            },
            "run_for_functions": [
                "microsoft_docs_search"
            ]
        }
    ]
}