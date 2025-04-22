// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { z } from "zod";

/**
 * Schema types supported by the fetcher
 */
export const SchemaTypeEnum = z.enum([
  "app_manifest",
  "declarative_agent_manifest",
  "api_plugin_manifest",
]);
export type SchemaType = z.infer<typeof SchemaTypeEnum>;

/**
 * Interface for schema repository information
 */
interface SchemaRepository {
  baseUrl: string;
}

/**
 * Repository configurations for each schema type
 */
const schemaRepositories: Record<SchemaType, SchemaRepository> = {
  app_manifest: {
    baseUrl: `https://developer.microsoft.com/json-schemas/teams/{{version}}/MicrosoftTeams.schema.json`,
  },
  declarative_agent_manifest: {
    baseUrl: `https://developer.microsoft.com/json-schemas/copilot/declarative-agent/{{version}}/schema.json`,
  },
  api_plugin_manifest: {
    baseUrl: `https://developer.microsoft.com/json-schemas/copilot/plugin/{{version}}/schema.json`,
  },
};

/**
 * Cache for previously fetched schemas
 */
const schemaCache: Record<string, any> = {};

/**
 * Fetch schema from the appropriate repository
 * @param schemaName The type of schema to fetch
 * @param schemaVersion The version of the schema to fetch
 * @returns The requested schema as a JSON object with schema_url and content properties
 */
export async function fetchSchema(schemaName: SchemaType, schemaVersion: string): Promise<string> {
  // Create a cache key
  const cacheKey = `${schemaName}:${schemaVersion}`;

  // Check if we have this schema in cache
  if (schemaCache[cacheKey]) {
    return schemaCache[cacheKey];
  }

  const repository = schemaRepositories[schemaName];
  if (!repository) {
    return `Unknown schema name: ${schemaName}`;
  }

  try {
    const url = repository.baseUrl.replace("{{version}}", schemaVersion);
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`HTTP error with status: ${response.status}`);
    }

    const content = await response.json();

    const result = JSON.stringify({
      schema_url: url,
      content: content,
    });

    // Save to cache
    schemaCache[cacheKey] = result;

    return result;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    return `Failed fetching schema at version: ${schemaVersion}, error: ${errorMessage}. Try different version.`;
  }
}
