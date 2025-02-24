{
    "id": "${{AAD_APP_OBJECT_ID}}",
    "appId": "${{AAD_APP_CLIENT_ID}}",
    "name": "{{appName}}-aad",
    "accessTokenAcceptedVersion": 2,
    "signInAudience": "AzureADMyOrg",
    "optionalClaims": {
        "idToken": [],
        "accessToken": [
            {
                "name": "idtyp",
                "source": null,
                "essential": false,
                "additionalProperties": []
            }
        ],
        "saml2Token": []
    },
    "oauth2Permissions": [
        {
            "adminConsentDescription": "Allows Copilot to read repair records on your behalf.",
            "adminConsentDisplayName": "Read repairs",
            "id": "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}",
            "isEnabled": true,
            "type": "User",
            "userConsentDescription": "Allows Copilot to read repair records.",
            "userConsentDisplayName": "Read repairs",
            "value": "repairs_read"
        }
    ],
{{#MicrosoftEntra}}
    "preAuthorizedApplications": [
        {
            "appId": "ab3be6b7-f5df-413d-ac2d-abf1e3fd9c0b",
            "permissionIds": [
                "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
            ]
        }
    ],
{{/MicrosoftEntra}}
    "replyUrlsWithType": [
        {
{{#MicrosoftEntra}}
           "url": "https://teams.microsoft.com/api/platform/v1.0/oAuthConsentRedirect",
{{/MicrosoftEntra}}
{{^MicrosoftEntra}}
           "url": "https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect",
{{/MicrosoftEntra}}
           "type": "Web"
        }
    ],
    "identifierUris": [
{{#MicrosoftEntra}}
        "api://${{OPENAPI_SERVER_DOMAIN}}/${{AAD_APP_CLIENT_ID}}",
        "${{AADAUTHCODE_APPLICATION_ID_URI}}"
{{/MicrosoftEntra}}
{{^MicrosoftEntra}}
        "api://${{AAD_APP_CLIENT_ID}}"
{{/MicrosoftEntra}}
    ]
}