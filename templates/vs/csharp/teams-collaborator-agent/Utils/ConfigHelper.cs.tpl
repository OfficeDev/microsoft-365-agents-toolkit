namespace {{SafeProjectName}}.Utils
{
    /// <summary>
    /// Configuration for AI models used by different capabilities
    /// </summary>
    public class ModelConfig
    {
        public string Model { get; set; } = string.Empty;
        public string ApiKey { get; set; } = string.Empty;
        public string Endpoint { get; set; } = string.Empty;
    }

    /// <summary>
    /// Utility class for managing AI model configurations and environment validation
    /// </summary>
    public static class ConfigHelper
    {
        private static ConfigOptions? _config;

        /// <summary>
        /// Initialize the configuration helper with ConfigOptions
        /// </summary>
        public static void Initialize(ConfigOptions config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));

            // Auto-configure database type based on Azure environment
            if (config.RunningOnAzure == "1" && string.IsNullOrEmpty(config.Database.Type))
            {
                config.Database.Type = "mssql";
            }
        }

        /// <summary>
        /// Get database configuration
        /// </summary>
        public static DatabaseConfigOptions GetDatabaseConfig()
        {
            if (_config == null)
            {
                throw new InvalidOperationException(
                    "ConfigHelper not initialized. Call Initialize() first."
                );
            }

            return _config.Database;
        }

        /// <summary>
        /// Model configurations for different capabilities
        /// </summary>
        public static class AIModels
        {
            /// <summary>
            /// Manager Capability - Uses lighter, faster model for routing decisions
            /// </summary>
            public static ModelConfig Manager
            {
                get
                {
                    if (_config == null)
                        throw new InvalidOperationException("ConfigHelper not initialized.");

                    return new ModelConfig
                    {
                        Model = "gpt-4o-mini",
{{#useAzureOpenAI}}
                        ApiKey = _config.Azure.OpenAIApiKey,
                        Endpoint = _config.Azure.OpenAIEndpoint,
{{/useAzureOpenAI}}
{{#useOpenAI}}
                        ApiKey = _config.OpenAI.ApiKey,
                        Endpoint = _config.OpenAI.Endpoint ?? string.Empty,
{{/useOpenAI}}
                    };
                }
            }

            /// <summary>
            /// Summarizer Capability - Uses more capable model for complex analysis
            /// </summary>
            public static ModelConfig Summarizer
            {
                get
                {
                    if (_config == null)
                        throw new InvalidOperationException("ConfigHelper not initialized.");

                    return new ModelConfig
                    {
{{#useAzureOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.Azure.OpenAIDeploymentName)
                            ? _config.Azure.OpenAIDeploymentName
                            : "gpt-4o",
                        ApiKey = _config.Azure.OpenAIApiKey,
                        Endpoint = _config.Azure.OpenAIEndpoint,
{{/useAzureOpenAI}}
{{#useOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.OpenAI.DefaultModel)
                            ? _config.OpenAI.DefaultModel
                            : "gpt-4o",
                        ApiKey = _config.OpenAI.ApiKey,
                        Endpoint = _config.OpenAI.Endpoint ?? string.Empty,
{{/useOpenAI}}
                    };
                }
            }

            /// <summary>
            /// Action Items Capability - Uses capable model for analysis and task management
            /// </summary>
            public static ModelConfig ActionItems
            {
                get
                {
                    if (_config == null)
                        throw new InvalidOperationException("ConfigHelper not initialized.");

                    return new ModelConfig
                    {
{{#useAzureOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.Azure.OpenAIDeploymentName)
                            ? _config.Azure.OpenAIDeploymentName
                            : "gpt-4o",
                        ApiKey = _config.Azure.OpenAIApiKey,
                        Endpoint = _config.Azure.OpenAIEndpoint,
{{/useAzureOpenAI}}
{{#useOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.OpenAI.DefaultModel)
                            ? _config.OpenAI.DefaultModel
                            : "gpt-4o",
                        ApiKey = _config.OpenAI.ApiKey,
                        Endpoint = _config.OpenAI.Endpoint ?? string.Empty,
{{/useOpenAI}}
                    };
                }
            }

            /// <summary>
            /// Search Capability - Uses capable model for semantic search and deep linking
            /// </summary>
            public static ModelConfig Search
            {
                get
                {
                    if (_config == null)
                        throw new InvalidOperationException("ConfigHelper not initialized.");

                    return new ModelConfig
                    {
{{#useAzureOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.Azure.OpenAIDeploymentName)
                            ? _config.Azure.OpenAIDeploymentName
                            : "gpt-4o",
                        ApiKey = _config.Azure.OpenAIApiKey,
                        Endpoint = _config.Azure.OpenAIEndpoint,
{{/useAzureOpenAI}}
{{#useOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.OpenAI.DefaultModel)
                            ? _config.OpenAI.DefaultModel
                            : "gpt-4o",
                        ApiKey = _config.OpenAI.ApiKey,
                        Endpoint = _config.OpenAI.Endpoint ?? string.Empty,
{{/useOpenAI}}
                    };
                }
            }

            /// <summary>
            /// Default model configuration (fallback)
            /// </summary>
            public static ModelConfig Default
            {
                get
                {
                    if (_config == null)
                        throw new InvalidOperationException("ConfigHelper not initialized.");

                    return new ModelConfig
                    {
{{#useAzureOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.Azure.OpenAIDeploymentName)
                            ? _config.Azure.OpenAIDeploymentName
                            : "gpt-4o",
                        ApiKey = _config.Azure.OpenAIApiKey,
                        Endpoint = _config.Azure.OpenAIEndpoint,
{{/useAzureOpenAI}}
{{#useOpenAI}}
                        Model = !string.IsNullOrEmpty(_config.OpenAI.DefaultModel)
                            ? _config.OpenAI.DefaultModel
                            : "gpt-4o",
                        ApiKey = _config.OpenAI.ApiKey,
                        Endpoint = _config.OpenAI.Endpoint ?? string.Empty,
{{/useOpenAI}}
                    };
                }
            }
        }

        /// <summary>
        /// Helper function to get model config for a specific capability
        /// </summary>
        public static ModelConfig GetModelConfig(string capabilityType)
        {
            return capabilityType.ToLower() switch
            {
                "manager" => AIModels.Manager,
                "summarizer" => AIModels.Summarizer,
                "actionitems" => AIModels.ActionItems,
                "search" => AIModels.Search,
                _ => AIModels.Default,
            };
        }

        /// <summary>
        /// Environment validation
        /// </summary>
        public static void ValidateEnvironment(ILogger logger)
        {
            if (_config == null)
            {
                throw new InvalidOperationException(
                    "ConfigHelper not initialized. Call Initialize() first."
                );
            }

            var hasModelConfig = false;

{{#useAzureOpenAI}}
            var azureApiKey = _config.Azure.OpenAIApiKey;
            var azureEndpoint = _config.Azure.OpenAIEndpoint;
            hasModelConfig =
                !string.IsNullOrEmpty(azureApiKey) && !string.IsNullOrEmpty(azureEndpoint);

            if (hasModelConfig)
            {
                logger.LogDebug("✓ Using Azure OpenAI configuration from appsettings");
            }
{{/useAzureOpenAI}}
{{#useOpenAI}}
            var openAiApiKey = _config.OpenAI.ApiKey;
            hasModelConfig = !string.IsNullOrEmpty(openAiApiKey);

            if (hasModelConfig)
            {
                logger.LogDebug("✓ Using OpenAI configuration");
            }
{{/useOpenAI}}

            if (!hasModelConfig)
            {
{{#useAzureOpenAI}}
                throw new InvalidOperationException(
                    "Missing required Azure OpenAI configuration. Please provide Azure:OpenAIApiKey and Azure:OpenAIEndpoint in appsettings.json"
                );
{{/useAzureOpenAI}}
{{#useOpenAI}}
                throw new InvalidOperationException(
                    "Missing required OpenAI configuration. Please provide OpenAI:ApiKey in appsettings.json"
                );
{{/useOpenAI}}
            }

            if (_config.Database.Type == "mssql")
            {
                if (string.IsNullOrEmpty(_config.Database.ConnectionString))
                {
                    logger.LogWarning(
                        "SQL Server configuration incomplete. Missing: SQL_CONNECTION_STRING. Falling back to SQLite."
                    );
                    _config.Database.Type = "sqlite";
                }
                else
                {
                    logger.LogDebug("✓ SQL Server configuration validated");
                }
            }

            logger.LogDebug("💾 Using database: {DatabaseType}", _config.Database.Type);
            logger.LogDebug("✓ Environment validation passed");
        }

        /// <summary>
        /// Model configuration logging
        /// </summary>
        public static void LogModelConfigs(ILogger logger)
        {
            logger.LogDebug("🤖 AI Model Configuration:");
            logger.LogDebug("  Manager Capability: {Model}", AIModels.Manager.Model);
            logger.LogDebug("  Summarizer Capability: {Model}", AIModels.Summarizer.Model);
            logger.LogDebug("  Action Items Capability: {Model}", AIModels.ActionItems.Model);
            logger.LogDebug("  Search Capability: {Model}", AIModels.Search.Model);
            logger.LogDebug("  Default Model: {Model}", AIModels.Default.Model);
        }

        /// <summary>
        /// Check if running on Azure
        /// </summary>
        public static bool IsRunningOnAzure()
        {
            return _config?.RunningOnAzure == "1";
        }
    }
}
