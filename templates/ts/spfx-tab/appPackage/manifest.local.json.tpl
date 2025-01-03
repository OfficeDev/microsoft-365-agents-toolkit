{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
    "manifestVersion": "1.19",
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
            "contentUrl": "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/TeamsWorkBench.aspx%3Fteams%26personal%26componentId={{componentId}}%26forceLocale={locale}%26loadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js",
            "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
            "scopes": [
                "personal"
            ]
        }
    ],
    "configurableTabs": [
        {
            "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/TeamsWorkBench.aspx%3FcomponentId={{componentId}}%26openPropertyPane=true%26teams%26forceLocale={locale}%26loadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js",
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