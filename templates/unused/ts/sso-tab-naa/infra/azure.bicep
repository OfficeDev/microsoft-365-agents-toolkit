param resourceBaseName string
param aadAppClientId string
param aadAppTenantId string
param aadAppOauthAuthorityHost string
param location string = resourceGroup().location
param staticWebAppName string = resourceBaseName
param staticWebAppSku string
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

// Azure Static Web Apps that hosts your static web site
resource swa 'Microsoft.Web/staticSites@2022-09-01' = {
  name: staticWebAppName
  // SWA do not need location setting
  location: 'centralus'
  sku: {
    name: staticWebAppSku
    tier: staticWebAppSku
  }
  properties:{}
}

var siteDomain = swa.properties.defaultHostname
var tabEndpoint = 'https://${siteDomain}'
var aadApplicationIdUri = 'api://${siteDomain}/${aadAppClientId}'
var oauthAuthority = uri(aadAppOauthAuthorityHost, aadAppTenantId)

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output TAB_DOMAIN string = siteDomain
output TAB_HOSTNAME string = siteDomain
output TAB_ENDPOINT string = 'https://${siteDomain}'
output AZURE_STATIC_WEB_APPS_RESOURCE_ID string = swa.id
