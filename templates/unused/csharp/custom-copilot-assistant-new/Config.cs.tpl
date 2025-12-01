namespace {{SafeProjectName}}
{
    public class ConfigOptions
    {
        public string BOT_ID { get; set; }
        public string BOT_PASSWORD { get; set; }
        public string BOT_TYPE { get; set; }
        public string BOT_TENANT_ID { get; set; }
{{#useOpenAI}}
        public OpenAIConfigOptions OpenAI { get; set; }
{{/useOpenAI}}
{{#useAzureOpenAI}}
        public AzureConfigOptions Azure { get; set; }
{{/useAzureOpenAI}}
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
