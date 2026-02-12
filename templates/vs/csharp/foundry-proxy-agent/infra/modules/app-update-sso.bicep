// Application Update Module for SSO
// Updates an existing App Registration with SSO configuration
// This module adds: redirect URIs, API scopes, pre-authorized apps, federated credentials

extension microsoftGraphV1

@description('The App ID (Client ID) of the existing application registration')
param botId string

@description('The Object ID of the existing application registration')
param botObjectId string

@description('Tenant ID where the application is registered')
param tenantId string

@description('Location for resources')
param location string = resourceGroup().location

// Reference the existing application by Object ID
resource existingApp 'Microsoft.Graph/applications@v1.0' existing = {
  uniqueName: botObjectId
}

// Update the application with SSO configuration
resource updatedApplication 'Microsoft.Graph/applications@v1.0' = {
  uniqueName: existingApp.uniqueName
  displayName: existingApp.displayName
  signInAudience: 'AzureADMyOrg'
  identifierUris: [
    'api://botid-${botId}'
  ]
  web: {
    redirectUris: [
      'https://token.botframework.com/.auth/web/redirect'
    ]
    implicitGrantSettings: {
      enableIdTokenIssuance: false
      enableAccessTokenIssuance: false
    }
  }
  
  api: {
    requestedAccessTokenVersion: 2
    oauth2PermissionScopes: [
      {
        id: guid(botId, 'access_as_user')
        adminConsentDescription: 'Default scope for Agent SSO access'
        adminConsentDisplayName: 'Agent SSO'
        userConsentDescription: 'Default scope for Agent SSO access'
        userConsentDisplayName: 'Agent SSO'
        value: 'access_as_user'
        type: 'User'
        isEnabled: true
      }
    ]
    preAuthorizedApplications: [
      {
        // Teams web client
        appId: '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
      {
        // Teams desktop client
        appId: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
      {
        // Microsoft 365 web application
        appId: '4765445b-32c6-49b0-83e6-1d93765276ca'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
      {
        // Microsoft 365 desktop application
        appId: '0ec893e0-5785-4de6-99da-4ed124e5296c'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
      {
        // Microsoft 365 mobile application Outlook desktop application
        appId: 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
      {
        // Outlook web application
        appId: 'bc59ab01-8403-45c6-8796-ac3ef710b3e3'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
      {
        // Outlook mobile application
        appId: '27922004-5251-4030-b22d-91ecd9a37ea4'
        delegatedPermissionIds: [
          guid(botId, 'access_as_user')
        ]
      }
    ]
  }
  
  requiredResourceAccess: [
    {
      // OpenID permissions & offline_access
      resourceAppId: '00000003-0000-0000-c000-000000000000'
      resourceAccess: [
        {
          // openid
          id: '37f7f235-527c-4136-accd-4a02d197296e'
          type: 'Scope'
        }
        {
          // profile
          id: '14dad69e-099b-42c9-810b-d002981feec1'
          type: 'Scope'
        }
        {
          // email
          id: '64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0'
          type: 'Scope'
        }
        {
          // offline_access
          id: '7427e0e9-2fba-42fe-b0c0-848c9e6a8182'
          type: 'Scope'
        }
      ]
    }
    {
      // Azure Machine Learning Services
      // Required for Microsoft Foundry agent SSO
      resourceAppId: '18a66f5f-dbdf-4c17-9dd7-1634712a9cbe'
      resourceAccess: [
        {
          // user_impersonation
          id: '1a7925b5-f871-417a-9b8b-303f9f29fa10'
          type: 'Scope'
        }
      ]
    }
  ]
}

// Call API to encode tenant ID
module tenantIdEncoder 'guid-encoder.bicep' = {
  name: 'encode-tenant-${uniqueString(tenantId)}'
  params: {
    guidToEncode: tenantId
    location: location
  }
}

var calculatedEncodedTenantId = tenantIdEncoder.outputs.encodedGuid

// Construct federated credential subject
var myfciSubject = '/eid1/c/pub/t/${calculatedEncodedTenantId}/a/9ExAW52n_ky4ZiS_jhpJIQ/${guid(botId, 'BotServiceOauthConnection')}'

// Add Federated Identity Credential for Azure Bot Service token exchange
resource federatedCredential 'Microsoft.Graph/applications/federatedIdentityCredentials@v1.0' = {
  name: '${updatedApplication.uniqueName}/${guid(botId, 'BotServiceOauthConnection')}'
  audiences: [
    'api://AzureADTokenExchange'
  ]
  issuer: '${environment().authentication.loginEndpoint}${tenantId}/v2.0'
  subject: myfciSubject
  description: 'Federated credential for Azure Bot Service token exchange'
}

// Outputs
output aadAppId string = botId
output aadAppObjectId string = botObjectId
output aadAppIdUri string = 'api://botid-${botId}'
output servicePrincipalId string = existingApp.id
output fciName string = federatedCredential.name
output fciSubject string = myfciSubject
