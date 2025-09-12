const { TokenCredentials, ManagedIdentityCredential } = require('@azure/identity');

function createConfig() {
  const clientId = process.env.CLIENT_ID;
  const tenantId = process.env.TENANT_ID;
  const botType = process.env.BOT_TYPE;

  if (!clientId) {
    throw new Error('CLIENT_ID environment variable is required');
  }

  let credentials;

  if (botType === 'UserAssignedMsi') {
    credentials = new ManagedIdentityCredential(clientId);
  } else {
    credentials = new TokenCredentials(clientId);
  }

  return {
    clientId,
    tenantId,
    botType,
    credentials
  };
}

module.exports.createConfig = createConfig;