{
    "id": "${{AAD_APP_OBJECT_ID}}",
    "appId": "${{AAD_APP_CLIENT_ID}}",
    "displayName": "{{appName}}-aad",
    "identifierUris": [
        "api://${{OPENAPI_SERVER_DOMAIN}}/${{AAD_APP_CLIENT_ID}}"
    ],
    "signInAudience": "AzureADMyOrg",
    "api": {
        "requestedAccessTokenVersion": 2,
        "oauth2PermissionScopes": [
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
        "preAuthorizedApplications": [
            {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "00000002-0000-0ff1-ce00-000000000000",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "bc59ab01-8403-45c6-8796-ac3ef710b3e3",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "0ec893e0-5785-4de6-99da-4ed124e5296c",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "4765445b-32c6-49b0-83e6-1d93765276ca",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "4345a7b9-9a63-4910-a426-35363201d503",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            },
            {
                "appId": "27922004-5251-4030-b22d-91ecd9a37ea4",
                "delegatedPermissionIds": [
                    "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
                ]
            }
        ]
    },
    "info": {},
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
    "publicClient": {},
    "web": {
        "implicitGrantSettings": {}
    },
    "spa": {}
}
