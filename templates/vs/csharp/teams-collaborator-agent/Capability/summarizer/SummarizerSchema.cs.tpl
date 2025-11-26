namespace {{SafeProjectName}}.Capability.Summarizer
{
    /// <summary>
    /// Function schemas for the summarizer capability
    /// Not needed for this capability because it has no custom functions with parameters
    /// </summary>
    public static class SummarizerSchema
    {
        // The summarize_conversation function doesn't require a schema
        // because it takes no parameters (async () => {...})
        // The function directly accesses context.Memory to retrieve messages
    }
}
