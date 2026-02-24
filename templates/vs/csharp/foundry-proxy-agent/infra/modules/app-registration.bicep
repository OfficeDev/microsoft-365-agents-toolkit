// Application Registration Module - Federated Credentials & Service Principal
// The SSO app is created by aadApp/create + aadApp/update YAML actions
// This module only creates: Federated Identity Credential + Service Principal
// This avoids Bicep PUT on Microsoft.Graph/applications which requires serviceManagementReference

extension microsoftGraphV1

@description('BotID this should match the Microsoft App ID in the Azure Bot Service Configuration')
param botId string

@description('Tenant ID where the application will be registered')
param tenantId string

@description('Pre-encoded tenant ID in Base64URL format (from guid-encoder module)')
param encodedTenantId string

@description('SSO App display name for service principal')
param ssoAppName string

// Reference the existing SSO app created by aadApp/create + aadApp/update
// The aadApp/create action now sets uniqueName to enable Bicep reference
resource existingSsoApp 'Microsoft.Graph/applications@v1.0' existing = {
  uniqueName: ssoAppName
}

// Construct federated credential subject using pre-encoded tenant ID
// appId encode value is the Bot Service one. it is hardcoded on purpose.
var myfciSubject = '/eid1/c/pub/t/${encodedTenantId}/a/9ExAW52n_ky4ZiS_jhpJIQ/${guid(ssoAppName, 'BotServiceOauthConnection')}'

// Federated Identity Credential for Bot Service token exchange
// The name must be in format: <parent-resource-symbolic-name>/<child-resource-name>
resource federatedCredential 'Microsoft.Graph/applications/federatedIdentityCredentials@v1.0' = {
  name: '${existingSsoApp.uniqueName}/${guid(ssoAppName, 'BotServiceOauthConnection')}'
  audiences: [
    'api://AzureADTokenExchange'
  ]
  issuer: '${environment().authentication.loginEndpoint}${tenantId}/v2.0'
  subject: myfciSubject
  description: 'Federated credential for Bot Service token exchange'
}

// Service Principal for the SSO application
resource ssoServicePrincipal 'Microsoft.Graph/servicePrincipals@v1.0' = {
  appId: existingSsoApp.appId
  accountEnabled: true
  displayName: ssoAppName
  servicePrincipalType: 'Application'
  tags: [
    'WindowsAzureActiveDirectoryIntegratedApp'
  ]
}

// Outputs for other modules
output aadAppId string = existingSsoApp.appId
output aadAppObjectId string = existingSsoApp.id
output aadAppIdUri string = 'api://botid-${botId}'
output servicePrincipalId string = ssoServicePrincipal.id
output servicePrincipalObjectId string = ssoServicePrincipal.id
output fciName string = federatedCredential.name
output fciSubject string = myfciSubject
