namespace {{SafeProjectName}}
{
    public class ConfigOptions
    {
        public TeamsConfigOptions Teams { get; set; } = new();
{{#useOpenAI}}
        public OpenAIConfigOptions OpenAI { get; set; } = new();
{{/useOpenAI}}
{{#useAzureOpenAI}}
        public AzureConfigOptions Azure { get; set; } = new();
{{/useAzureOpenAI}}
    }

    public class TeamsConfigOptions
    {
        public string? BotType { get; set; }
        public string? ClientId { get; set; }
        public string? ClientSecret { get; set; }
        public string? TenantId { get; set; }
    }
{{#useOpenAI}}
    /// <summary>
    /// Options for Open AI
    /// </summary>
    public class OpenAIConfigOptions
    {
        public string ApiKey { get; set; } = string.Empty;
        public string DefaultModel = "gpt-3.5-turbo";
    }
{{/useOpenAI}}
{{#useAzureOpenAI}}
    /// <summary>
    /// Options for Azure OpenAI and Azure Content Safety
    /// </summary>
    public class AzureConfigOptions
    {
        public string OpenAIApiKey { get; set; } = string.Empty;
        public string OpenAIEndpoint { get; set; } = string.Empty;
        public string OpenAIDeploymentName { get; set; } = string.Empty;
    }
{{/useAzureOpenAI}}
}
