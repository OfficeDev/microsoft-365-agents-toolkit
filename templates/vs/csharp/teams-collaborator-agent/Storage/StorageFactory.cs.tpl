using {{SafeProjectName}}.Utils;

namespace {{SafeProjectName}}.Storage
{
    public static class StorageFactory
    {
        public static async Task<IDatabase> CreateStorageAsync(
            ILogger logger,
            DatabaseConfigOptions? config = null
        )
        {
            var dbConfig = config ?? ConfigHelper.GetDatabaseConfig();

            if (dbConfig.Type == "mssql")
            {
                try
                {
                    logger.LogDebug("Initializing MSSQL storage");
                    var storage = new MssqlDatabase(logger, dbConfig);
                    await storage.InitializeAsync();
                    logger.LogDebug("MSSQL storage initialized");
                    return storage;
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Failed to initialize MSSQL, falling back to SQLite");
                }
            }

            logger.LogDebug("Initializing SQLite storage");
            var sqliteStorage = new SqliteDatabase(logger, dbConfig.SqlitePath);
            await sqliteStorage.InitializeAsync();
            logger.LogDebug("SQLite storage initialized");

            return sqliteStorage;
        }
    }
}
