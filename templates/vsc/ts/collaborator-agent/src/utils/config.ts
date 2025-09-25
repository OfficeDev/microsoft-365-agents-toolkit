import { ILogger } from "@microsoft/teams.common";

// Configuration for AI models used by different capabilities
export interface ModelConfig {
  model: string;
  apiKey: string;
  endpoint: string;
  apiVersion: string;
}

// Model configurations for different capabilities
export const AI_MODELS = {
  // Manager Capability - Uses lighter, faster model for routing decisions
  MANAGER: {
    model: "gpt-4o-mini",
    apiKey: process.env.AOAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Summarizer Capability - Uses more capable model for complex analysis
  SUMMARIZER: {
    model: process.env.AOAI_MODEL || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Action Items Capability - Uses capable model for analysis and task management
  ACTION_ITEMS: {
    model: process.env.AOAI_MODEL || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Search Capability - Uses capable model for semantic search and deep linking
  SEARCH: {
    model: process.env.AOAI_MODEL || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,

  // Default model configuration (fallback)
  DEFAULT: {
    model: process.env.AOAI_MODEL || "gpt-4o",
    apiKey: process.env.AOAI_API_KEY!,
    endpoint: process.env.AOAI_ENDPOINT!,
    apiVersion: "2025-04-01-preview",
  } as ModelConfig,
};

// Helper function to get model config for a specific capability
export function getModelConfig(capabilityType: string): ModelConfig {
  switch (capabilityType.toLowerCase()) {
    case "manager":
      return AI_MODELS.MANAGER;
    case "summarizer":
      return AI_MODELS.SUMMARIZER;
    case "actionitems":
      return AI_MODELS.ACTION_ITEMS;
    case "search":
      return AI_MODELS.SEARCH;
    default:
      return AI_MODELS.DEFAULT;
  }
}

// Environment validation
export function validateEnvironment(logger: ILogger): void {
  const requiredEnvVars = ["AOAI_API_KEY", "AOAI_ENDPOINT"];
  const missing = requiredEnvVars.filter((envVar) => !process.env[envVar]);

  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(", ")}`);
  }
  logger.debug("✅ Environment validation passed");
}

// Model configuration logging
export function logModelConfigs(logger: ILogger): void {
  logger.debug("🔧 AI Model Configuration:");
  logger.debug(`  Manager Capability: ${AI_MODELS.MANAGER.model}`);
  logger.debug(`  Summarizer Capability: ${AI_MODELS.SUMMARIZER.model}`);
  logger.debug(`  Action Items Capability: ${AI_MODELS.ACTION_ITEMS.model}`);
  logger.debug(`  Search Capability: ${AI_MODELS.SEARCH.model}`);
  logger.debug(`  Default Model: ${AI_MODELS.DEFAULT.model}`);
}
