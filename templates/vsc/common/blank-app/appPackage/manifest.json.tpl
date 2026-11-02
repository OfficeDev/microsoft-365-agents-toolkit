{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.27/MicrosoftTeams.schema.json",
    "manifestVersion": "1.27",
    "version": "1.0.0",
    {{#NewDeveloperPortalApis}}
    "id": "${{APP_ID}}",
    {{/NewDeveloperPortalApis}}
    {{^NewDeveloperPortalApis}}
    "id": "${{TEAMS_APP_ID}}",
    {{/NewDeveloperPortalApis}}
    "developer": {
        "name": "My App, Inc.",
        "websiteUrl": "https://www.example.com",
        "privacyUrl": "https://www.example.com/privacy",
        "termsOfUseUrl": "https://www.example.com/termofuse"
    },
    "icons": {
        "color": "color.png",
        "outline": "outline.png"
    },
    "name": {
        "short": "{{appName}}${{APP_NAME_SUFFIX}}",
        "full": "Full name for {{appName}}"
    },
    "description": {
        "short": "Short description for {{appName}}",
        "full": "Full description for {{appName}}"
    },
    "accentColor": "#FFFFFF",
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": []
}
