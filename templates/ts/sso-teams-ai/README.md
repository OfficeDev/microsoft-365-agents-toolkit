# Teams Toolkit Configuration: Oauth

Use this if you want to enable user authentication in your Teams application.

## How to update scopes

1. In the `aad.manifest.json` file, update the `requiredResourceAccess` list to add the required scopes.

2. In the `infra/botRegistration/azurebot.bicep` file, under the `botServicesMicrosoftGraphConnection` resource, update the `properties.scopes` string to be a comma-delimeted list of the required scopes.

### Example

If you want to add the `People.Read.All` and `User.ReadBasic.All` scopes.

1. Your `requiredResourceAccess` property should look like:

```json
"requiredResourceAccess": [
    {
        "resourceAppId": "Microsoft Graph",
        "resourceAccess": [
            {
                "id": "People.Read.All",
                "type": "Scope"
            }
        ]
    },
    {
        "resourceAppId": "Microsoft Graph",
        "resourceAccess": [
            {
                "id": "User.ReadBasic.All",
                "type": "Scope"
            }
        ]
    },
]
```

2. Update the `properties.scopes` to be `People.Read.All,User.ReadBasic.All`.
