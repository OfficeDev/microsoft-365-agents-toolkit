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
    /// Options for Open AI
    /// </summary>
    public class OpenAIConfigOptions
    {
        public string ApiKey { get; set; }
        public string DefaultModel = "gpt-3.5-turbo";
    }
{{/useOpenAI}}
{{#useAzureOpenAI}}
    /// <summary>
    /// Options for Azure OpenAI and Azure Content Safety
    /// </summary>
    public class AzureConfigOptions
    {
        public string OpenAIApiKey { get; set; }
        public string OpenAIEndpoint { get; set; }
        public string OpenAIDeploymentName { get; set; }
    }
{{/useAzureOpenAI}}
}