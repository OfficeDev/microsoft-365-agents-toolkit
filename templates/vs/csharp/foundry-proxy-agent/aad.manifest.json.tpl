{
    "id": "${{SSO_APP_OBJECT_ID}}",
    "appId": "${{SSO_APP_ID}}",
    "displayName": "{{appName}}-${{RESOURCE_SUFFIX}}-${{APP_NAME_SUFFIX}}-UserAuth",
    "identifierUris": [
        "api://botid-${{BOT_ID}}"
    ],
    "signInAudience": "AzureADMyOrg",
    "api": {
        "requestedAccessTokenVersion": 2,
        "oauth2PermissionScopes": [
            {
                "adminConsentDescription": "Default scope for Agent SSO access",
                "adminConsentDisplayName": "Agent SSO",
                "id": "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": "Default scope for Agent SSO access",
                "userConsentDisplayName": "Agent SSO",
                "value": "access_as_user"
            }
        ],
        "preAuthorizedApplications": [
            {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "4765445b-32c6-49b0-83e6-1d93765276ca",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "0ec893e0-5785-4de6-99da-4ed124e5296c",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "bc59ab01-8403-45c6-8796-ac3ef710b3e3",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "27922004-5251-4030-b22d-91ecd9a37ea4",
                "delegatedPermissionIds": [
                    "${{SSO_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            }
        ]
    },
    "web": {
        "redirectUris": [
            "https://token.botframework.com/.auth/web/redirect"
        ],
        "implicitGrantSettings": {
            "enableIdTokenIssuance": false,
            "enableAccessTokenIssuance": false
        }
    },
    "requiredResourceAccess": [
        {
            "resourceAppId": "Microsoft Graph",
            "resourceAccess": [
                {
                    "id": "openid",
                    "type": "Scope"
                },
                {
                    "id": "profile",
                    "type": "Scope"
                },
                {
                    "id": "email",
                    "type": "Scope"
                },
                {
                    "id": "offline_access",
                    "type": "Scope"
                }
            ]
        },
        {
            "resourceAppId": "18a66f5f-dbdf-4c17-9dd7-1634712a9cbe",
            "resourceAccess": [
                {
                    "id": "1a7925b5-f871-417a-9b8b-303f9f29fa10",
                    "type": "Scope"
                }
            ]
        }
    ]
}
