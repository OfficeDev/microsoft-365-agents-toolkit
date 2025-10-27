using Microsoft.Extensions.Logging;

namespace {{SafeProjectName}}.Storage;

public static class StorageFactory
{
    public static IConversationStorage CreateStorage(StorageConfigOptions config, ILogger logger)
    {
        logger.LogInformation($"Creating storage of type: {config.Type}");

        return config.Type.ToLowerInvariant() switch
        {
            "sqlite" => new SqliteConversationStorage(config.ConnectionString),
            "sqlserver" => new SqlServerConversationStorage(config.ConnectionString),
            _ => throw new InvalidOperationException($"Unsupported storage type: {config.Type}")
        };
    }
}
