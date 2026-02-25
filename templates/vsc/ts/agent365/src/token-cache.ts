// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// ------------------------------------------------------------------------------

export function createAgenticTokenCacheKey(agentId: string, tenantId?: string): string {
  return tenantId ? `agentic-token-${agentId}-${tenantId}` : `agentic-token-${agentId}`;
}

// A simple example of custom token resolver which will be called by observability SDK when needing tokens for exporting telemetry
export const tokenResolver = (agentId: string, tenantId: string): string | null => {
  try {
    const cacheKey = createAgenticTokenCacheKey(agentId, tenantId);
    const cachedToken = tokenCache.get(cacheKey);
    return cachedToken ?? null;
  } catch (error) {
    console.error(`❌ Error resolving token for agent ${agentId}, tenant ${tenantId}:`, error);
    return null;
  }
};

class TokenCache {
  private cache = new Map<string, string>();
  set(key: string, token: string): void {
    this.cache.set(key, token);
    console.log(`🔐 Token cached for key: ${key}`);
  }
  get(key: string): string | null {
    const entry = this.cache.get(key);
    if (!entry) {
      console.log(`🔍 Token cache miss for key: ${key}`);
      return null;
    }
    return entry;
  }
  has(key: string): boolean {
    return this.cache.has(key);
  }
}

const tokenCache = new TokenCache();
export default tokenCache;