namespace {{SafeProjectName}}.Capability.ActionItems
{
    /// <summary>
    /// Function schemas for the action items capability
    /// Not needed for this capability because it has no custom functions with parameters
    /// </summary>
    public static class ActionItemsSchema
    {
        // The generate_action_items function doesn't require a schema
        // because it takes no parameters (async () => {...})
        // The function directly accesses context.Memory to retrieve messages
    }
}
