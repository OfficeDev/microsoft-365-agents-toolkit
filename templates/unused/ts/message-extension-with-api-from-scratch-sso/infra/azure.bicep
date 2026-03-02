@maxLength(20)
@minLength(4)
param resourceBaseName string
param functionAppSKU string
param aadAppClientId string
param aadAppTenantId string
param aadAppOauthAuthorityHost string
param location string = resourceGroup().location
param serverfarmsName string = resourceBaseName
param functionAppName string = resourceBaseName
var teamsMobileOrDesktopAppClientId = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
var teamsWebAppClientId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'
var officeWebAppClientId1 = '4345a7b9-9a63-4910-a426-35363201d503'
var officeWebAppClientId2 = '4765445b-32c6-49b0-83e6-1d93765276ca'
var outlookDesktopAppClientId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
var outlookWebAppClientId = '00000002-0000-0ff1-ce00-000000000000'
var officeUwpPwaClientId = '0ec893e0-5785-4de6-99da-4ed124e5296c'
var outlookOnlineAddInAppClientId = 'bc59ab01-8403-45c6-8796-ac3ef710b3e3'
var outlookMobileAppClientId = '27922004-5251-4030-b22d-91ecd9a37ea4'
var allowedClientApplications = '"${teamsMobileOrDesktopAppClientId}","${teamsWebAppClientId}","${officeWebAppClientId1}","${officeWebAppClientId2}","${outlookDesktopAppClientId}","${outlookWebAppClientId}","${officeUwpPwaClientId}","${outlookOnlineAddInAppClientId}","${outlookMobileAppClientId}"'

// Compute resources for Azure Functions
resource serverfarms 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: serverfarmsName
  location: location
  sku: {
    name: functionAppSKU // You can follow https://aka.ms/teamsfx-bicep-add-param-tutorial to add functionServerfarmsSku property to provisionParameters to override the default value "Y1".
  }
  properties: {}
}

// Azure Functions that hosts your function code
resource functionApp 'Microsoft.Web/sites@2021-02-01' = {
  name: functionAppName
  kind: 'functionapp'
  location: location
  properties: {
    serverFarmId: serverfarms.id
    httpsOnly: true
    siteConfig: {
      appSettings: [
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4' // Use Azure Functions runtime v4
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node' // Set runtime to NodeJS
        }
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure Functions from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~18' // Set NodeJS version to 18.x
        }
        {
          name: 'aadAppClientId'
          value: aadAppClientId
        }
        {
          name: 'aadAppTenantId'
          value: aadAppTenantId
        }
        {
          name: 'M365_AUTHORITY_HOST'
          value: aadAppOauthAuthorityHost
        }
        {
          name: 'WEBSITE_AUTH_AAD_ACL'
          value: '{"allowed_client_applications": [${allowedClientApplications}]}'
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
}
var apiEndpoint = 'https://${functionApp.properties.defaultHostName}'
var oauthAuthority = uri(aadAppOauthAuthorityHost, aadAppTenantId)
var aadApplicationIdUri = 'api://${functionApp.properties.defaultHostName}/${aadAppClientId}'

// Configure Azure Functions to use Azure AD for authentication.
resource authSettings 'Microsoft.Web/sites/config@2021-02-01' = {
  parent: functionApp
  name: 'authsettingsV2'
  properties: {
    globalValidation: {
      requireAuthentication: true
      unauthenticatedClientAction: 'Return401'
    }

    identityProviders: {
      azureActiveDirectory: {
        enabled: true
        registration: {
          openIdIssuer: oauthAuthority
          clientId: aadAppClientId
        }
        validation: {
          allowedAudiences: [
            aadAppClientId
            aadApplicationIdUri
          ]
          defaultAuthorizationPolicy: {
            allowedApplications: [
              teamsMobileOrDesktopAppClientId
              teamsWebAppClientId
              officeWebAppClientId1
              officeWebAppClientId2
              outlookDesktopAppClientId
              outlookWebAppClientId
              officeUwpPwaClientId
              outlookOnlineAddInAppClientId
            ]
          }
        }
      }
    }
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output API_FUNCTION_ENDPOINT string = apiEndpoint
output API_FUNCTION_RESOURCE_ID string = functionApp.id
output OPENAPI_SERVER_URL string = apiEndpoint
output OPENAPI_SERVER_DOMAIN string = functionApp.properties.defaultHostName
