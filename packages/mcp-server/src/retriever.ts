// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";

export const ResourceTypeEnum = z.enum(["documents", "samples", "issues", "code"]);
export type ResourceType = z.infer<typeof ResourceTypeEnum>;

/**
 * Retrieves resources from the API service
 * @param resourceType The type of resource to retrieve
 * @param question The question to use for retrieval
 * @returns The retrieved resources or error message as string
 */
export async function retrieveResource(
  resourceType: ResourceType,
  question: string
): Promise<string> {
  try {
    const apiEndpoint =
      process.env.RETRIEVER_API_ENDPOINT ||
      Buffer.from(
        // eslint-disable-next-line no-secrets/no-secrets
        "aHR0cHM6Ly9hZmQtd20zZGg1amM2NzU1cy1wcm9kLWhrZndnYmJqYjVhN2hyYnUuYjAxLmF6dXJlZmQubmV0L3JldHJpZXZlcg==",
        "base64"
      ).toString("utf-8");

    const response = await fetch(apiEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        resource_type: resourceType,
        question: question,
      }),
    });

    if (!response.ok) {
      return `Fail to retrieve resource, ${response.status}: ${response.statusText}`;
    }

    const data = await response.json();
    return data.data;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
    return `Error retrieving resource: ${errorMessage}`;
  }
}
