@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string
{{#useOpenAI}}
@secure()
param openAIApiKey string
{{/useOpenAI}}
{{#useAzureOpenAI}}
@secure()
param azureOpenAIApiKey string

param azureOpenAIEndpoint string
param azureOpenAIDeploymentName string
{{/useAzureOpenAI}}

param webAppSKU string

@maxLength(42)
param botDisplayName string

@secure()
param sqlAdminLogin string = 'sqladmin'
@secure()
param sqlAdminPassword string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param identityName string = resourceBaseName
param sqlServerName string = '${resourceBaseName}-sqlserver'
param sqlDatabaseName string = '${resourceBaseName}-db'
param location string = resourceGroup().location

resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  location: location
  name: identityName
}

// SQL Server
resource sqlServer 'Microsoft.Sql/servers@2023-05-01-preview' = {
  name: sqlServerName
  location: location
  properties: {
    administratorLogin: sqlAdminLogin
    administratorLoginPassword: sqlAdminPassword
    version: '12.0'
  }
}

// SQL Database
resource sqlDatabase 'Microsoft.Sql/servers/databases@2023-05-01-preview' = {
  parent: sqlServer
  name: sqlDatabaseName
  location: location
  sku: {
    name: 'Basic'
    tier: 'Basic'
    capacity: 5
  }
  properties: {
    collation: 'SQL_Latin1_General_CP1_CI_AS'
    maxSizeBytes: 2147483648 // 2GB
  }
}

// Allow Azure services to access SQL Server
resource sqlFirewallRule 'Microsoft.Sql/servers/firewallRules@2023-05-01-preview' = {
  parent: sqlServer
  name: 'AllowAllWindowsAzureIps'
  properties: {
    startIpAddress: '0.0.0.0'
    endIpAddress: '0.0.0.0'
  }
}

// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
}

// Web App that hosts your agent
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'Teams__ClientId'
          value: identity.properties.clientId
        }
        {
          name: 'Teams__TenantId'
          value: identity.properties.tenantId
        }
        {
          name: 'Teams__BotType'
          value: 'UserAssignedMsi'
        }
{{#useOpenAI}}
        {
          name: 'OpenAI__ApiKey'
          value: openAIApiKey
        }
{{/useOpenAI}}
{{#useAzureOpenAI}}
        {
          name: 'Azure__OpenAIApiKey'
          value: azureOpenAIApiKey
        }
        {
          name: 'Azure__OpenAIEndpoint'
          value: azureOpenAIEndpoint
        }
        {
          name: 'Azure__OpenAIDeploymentName'
          value: azureOpenAIDeploymentName
        }
{{/useAzureOpenAI}}
        {
          name: 'Database__Type'
          value: 'mssql'
        }
        {
          name: 'Database__ConnectionString'
          value: 'Server=tcp:${sqlServer.properties.fullyQualifiedDomainName},1433;Initial Catalog=${sqlDatabaseName};Persist Security Info=False;User ID=${sqlAdminLogin};Password=${sqlAdminPassword};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;'
        }
        {
          name: 'Database__Server'
          value: sqlServer.properties.fullyQualifiedDomainName
        }
        {
          name: 'Database__Database'
          value: sqlDatabaseName
        }
        {
          name: 'Database__Username'
          value: sqlAdminLogin
        }
        {
          name: 'Database__Password'
          value: sqlAdminPassword
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    identityClientId: identity.properties.clientId
    identityResourceId: identity.id
    identityTenantId: identity.properties.tenantId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
output SQL_SERVER_FQDN string = sqlServer.properties.fullyQualifiedDomainName
output SQL_DATABASE_NAME string = sqlDatabaseName
