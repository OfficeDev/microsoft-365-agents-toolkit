import { TokenCredentials, ManagedIdentityCredential } from '@azure/identity';

export interface ConfigInterface {
  clientId: string;
  tenantId?: string;
  botType?: string;
  credentials: TokenCredentials | ManagedIdentityCredential;
}

export function createConfig(): ConfigInterface {
  const clientId = process.env.CLIENT_ID;
  const tenantId = process.env.TENANT_ID;
  const botType = process.env.BOT_TYPE;

  if (!clientId) {
    throw new Error('CLIENT_ID environment variable is required');
  }

  let credentials: TokenCredentials | ManagedIdentityCredential;

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