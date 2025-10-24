namespace {{SafeProjectName}}
{
    public class ConfigOptions
    {
        public TeamsConfigOptions Teams { get; set; }
{{#useOpenAI}}
        public OpenAIConfigOptions OpenAI { get; set; }
{{/useOpenAI}}
{{#useAzureOpenAI}}
        public AzureConfigOptions Azure { get; set; }
{{/useAzureOpenAI}}
        public StorageConfigOptions Storage { get; set; }
    }

    public class TeamsConfigOptions
    {
        public string BotType { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string TenantId { get; set; }
    }
{{#useOpenAI}}
    /// <summary>
    /// Options for OpenAI
    /// </summary>
    public class OpenAIConfigOptions
    {
        public string ApiKey { get; set; }
        public string DefaultModel { get; set; } = "gpt-4o";
    }
{{/useOpenAI}}
{{#useAzureOpenAI}}
    /// <summary>
    /// Options for Azure OpenAI
    /// </summary>
    public class AzureConfigOptions
    {
        public string OpenAIApiKey { get; set; }
        public string OpenAIEndpoint { get; set; }
        public string OpenAIDeploymentName { get; set; }
    }
{{/useAzureOpenAI}}
    /// <summary>
    /// Options for conversation storage
    /// </summary>
    public class StorageConfigOptions
    {
        public string Type { get; set; } = "sqlite";
        public string ConnectionString { get; set; } = "Data Source=conversations.db";
    }
}
