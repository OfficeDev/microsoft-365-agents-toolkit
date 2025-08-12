{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.22/MicrosoftTeams.schema.json",
    "manifestVersion": "1.22",
    "id": "${{TEAMS_APP_ID}}",
    "version": "1.0.0",
    "developer": {
        "name": "SPFx + Teams Dev",
        "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
        "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
        "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
    },
    "name": {
        "short": "{{appName}}${{APP_NAME_SUFFIX}}",
        "full": "Full name for {{appName}}"
    },
    "description": {
        "short": "Short description of {{appName}}",
        "full": "Full description of {{appName}}"
    },
    "icons": {
        "color": "color.png",
        "outline": "outline.png"
    },
    "accentColor": "#004578",
    "staticTabs": [
        {
            "entityId": "{{componentId}}",
            "name": "{{webpartName}}",
            "contentUrl": "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId={{componentId}}%26forceLocale={locale}",
            "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
            "scopes": [
                "personal"
            ]
        }
    ],
    "configurableTabs": [
        {
            "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId={{componentId}}%26forceLocale={locale}",
            "canUpdateConfiguration": true,
            "scopes": [
                "team"
            ]
        }
    ],
    "permissions": [
        "identity",
        "messageTeamMembers"
    ],
    "validDomains": [
        "*.login.microsoftonline.com",
        "*.sharepoint.com",
        "*.sharepoint-df.com",
        "spoppe-a.akamaihd.net",
        "spoprod-a.akamaihd.net",
        "resourceseng.blob.core.windows.net",
        "msft.spoppe.com"
    ],
    "webApplicationInfo": {
        "resource": "https://{teamSiteDomain}",
        "id": "00000003-0000-0ff1-ce00-000000000000"
    }
}