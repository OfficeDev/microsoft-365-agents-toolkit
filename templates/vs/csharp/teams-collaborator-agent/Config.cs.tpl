namespace {{SafeProjectName}}
{
    public class ConfigOptions
    {
        public TeamsConfigOptions Teams { get; set; } = new();
{{#useAzureOpenAI}}
        public AzureConfigOptions Azure { get; set; } = new();
{{/useAzureOpenAI}}
{{#useOpenAI}}
        public OpenAIConfigOptions OpenAI { get; set; } = new();
{{/useOpenAI}}
        public DatabaseConfigOptions Database { get; set; } = new();
        public string? RunningOnAzure { get; set; }
    }

    public class TeamsConfigOptions
    {
        public string BotType { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
        public string TenantId { get; set; } = string.Empty;
    }
{{#useAzureOpenAI}}
    /// <summary>
    /// Options for Azure OpenAI
    /// </summary>
    public class AzureConfigOptions
    {
        public string OpenAIApiKey { get; set; } = string.Empty;
        public string OpenAIEndpoint { get; set; } = string.Empty;
        public string OpenAIDeploymentName { get; set; } = string.Empty;
    }
{{/useAzureOpenAI}}
{{#useOpenAI}}
    /// <summary>
    /// Options for OpenAI
    /// </summary>
    public class OpenAIConfigOptions
    {
        public string ApiKey { get; set; } = string.Empty;
        public string DefaultModel { get; set; } = "gpt-4o";
        public string? Endpoint { get; set; }
    }
{{/useOpenAI}}

    /// <summary>
    /// Database configuration options
    /// </summary>
    public class DatabaseConfigOptions
    {
        public string Type { get; set; } = "sqlite";
        public string? ConnectionString { get; set; }
        public string? Server { get; set; }
        public string? Database { get; set; }
        public string? Username { get; set; }
        public string? Password { get; set; }
        public string? SqlitePath { get; set; }
    }
}
