// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { z } from "zod";

export const ResourceTypeEnum = z.enum(["documents", "samples", "issues"]);
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
    const token = process.env.GITHUB_PERSONAL_ACCESS_TOKEN;
    const apiEndpoint = process.env.RETRIEVER_API_ENDPOINT;

    if (!token) {
      return "GITHUB_PERSONAL_ACCESS_TOKEN is not set, set it in your MCP server configuration to use this tool.";
    }

    if (!apiEndpoint) {
      return "RETRIEVER_API_ENDPOINT is not set, set it in your MCP server configuration to use this tool.";
    }

    const response = await fetch(apiEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`,
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
