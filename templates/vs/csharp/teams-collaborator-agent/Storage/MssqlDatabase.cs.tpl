using System.Data;
using System.Text.Json;
using {{SafeProjectName}}.Utils;
using Microsoft.Data.SqlClient;

namespace {{SafeProjectName}}.Storage
{
    public class MssqlDatabase : IDatabase
    {
        private readonly ILogger _logger;
        private readonly DatabaseConfigOptions _config;
        private SqlConnection? _connection;
        private bool _isInitialized;

        public MssqlDatabase(ILogger logger, DatabaseConfigOptions config)
        {
            _logger = logger;
            _config = config;
        }

        public async Task InitializeAsync()
        {
            if (_isInitialized)
            {
                return;
            }

            try
            {
                var connectionString = ResolveConnectionString();

                _connection = new SqlConnection(connectionString);
                await _connection.OpenAsync();

                await InitializeDatabaseAsync();
                _isInitialized = true;

                _logger.LogDebug("Connected to MSSQL database");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error connecting to MSSQL database");
                throw;
            }
        }

        private string ResolveConnectionString()
        {
            if (!string.IsNullOrEmpty(_config.ConnectionString))
            {
                return _config.ConnectionString;
            }

            var builder = new SqlConnectionStringBuilder
            {
                DataSource = _config.Server,
                InitialCatalog = _config.Database,
                UserID = _config.Username,
                Password = _config.Password,
                Encrypt = true,
                TrustServerCertificate = false,
            };

            return builder.ConnectionString;
        }

        private async Task InitializeDatabaseAsync()
        {
            if (_connection == null)
            {
                throw new InvalidOperationException("Database not connected");
            }

            try
            {
                await using var command = _connection.CreateCommand();

                command.CommandText =
                    @"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='conversations' AND xtype='U')
BEGIN
    CREATE TABLE conversations (
        id INT IDENTITY(1,1) PRIMARY KEY,
        conversation_id NVARCHAR(255) NOT NULL,
        role NVARCHAR(50) NOT NULL,
        name NVARCHAR(255) NOT NULL,
        content NVARCHAR(MAX) NOT NULL,
        activity_id NVARCHAR(255) NOT NULL,
        timestamp NVARCHAR(50) NOT NULL,
        blob NVARCHAR(MAX) NOT NULL
    )
END";
                await command.ExecuteNonQueryAsync();

                command.CommandText =
                    @"IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='idx_conversation_id' AND object_id = OBJECT_ID('conversations'))
BEGIN
    CREATE INDEX idx_conversation_id ON conversations(conversation_id)
END";
                await command.ExecuteNonQueryAsync();

                command.CommandText =
                    @"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='feedback' AND xtype='U')
BEGIN
    CREATE TABLE feedback (
        id INT IDENTITY(1,1) PRIMARY KEY,
        reply_to_id NVARCHAR(255) NOT NULL,
        reaction NVARCHAR(50) NOT NULL CHECK (reaction IN ('like','dislike')),
        feedback NVARCHAR(MAX),
        created_at DATETIME NOT NULL DEFAULT GETDATE()
    )
END";
                await command.ExecuteNonQueryAsync();

                _logger.LogDebug("MSSQL database tables initialized");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error initializing MSSQL tables");
                throw;
            }
        }

        public async Task ClearAllAsync()
        {
            EnsureConnection();

            try
            {
                await using var command = _connection!.CreateCommand();
                command.CommandText = "DELETE FROM conversations";
                await command.ExecuteNonQueryAsync();

                _logger.LogDebug("Cleared all conversations from MSSQL store");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error clearing all conversations");
                throw;
            }
        }

        public async Task<List<MessageRecord>> GetAsync(string conversationId)
        {
            EnsureConnection();

            try
            {
                var messages = new List<MessageRecord>();

                await using var command = _connection!.CreateCommand();
                command.CommandText =
                    "SELECT blob FROM conversations WHERE conversation_id = @conversationId ORDER BY timestamp ASC";
                command.Parameters.Add(
                    new SqlParameter("@conversationId", SqlDbType.NVarChar) { Value = conversationId }
                );

                await using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var blob = reader.GetString(0);
                    var message = JsonSerializer.Deserialize<MessageRecord>(blob);
                    if (message != null)
                    {
                        messages.Add(message);
                    }
                }

                return messages;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving messages");
                return new List<MessageRecord>();
            }
        }

        public async Task<List<MessageRecord>> GetMessagesByTimeRangeAsync(
            string conversationId,
            string startTime,
            string endTime
        )
        {
            EnsureConnection();

            try
            {
                var messages = new List<MessageRecord>();

                await using var command = _connection!.CreateCommand();
                command.CommandText =
                    @"SELECT blob FROM conversations
WHERE conversation_id = @conversationId
    AND timestamp >= @startTime
    AND timestamp <= @endTime
ORDER BY timestamp ASC";

                command.Parameters.Add(
                    new SqlParameter("@conversationId", SqlDbType.NVarChar) { Value = conversationId }
                );
                command.Parameters.Add(
                    new SqlParameter("@startTime", SqlDbType.NVarChar) { Value = startTime }
                );
                command.Parameters.Add(
                    new SqlParameter("@endTime", SqlDbType.NVarChar) { Value = endTime }
                );

                await using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var blob = reader.GetString(0);
                    var message = JsonSerializer.Deserialize<MessageRecord>(blob);
                    if (message != null)
                    {
                        messages.Add(message);
                    }
                }

                return messages;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving messages by time range");
                return new List<MessageRecord>();
            }
        }

        public async Task<List<MessageRecord>> GetRecentMessagesAsync(
            string conversationId,
            int limit = 10
        )
        {
            var messages = await GetAsync(conversationId);
            return messages.TakeLast(limit).ToList();
        }

        public async Task ClearConversationAsync(string conversationId)
        {
            EnsureConnection();

            try
            {
                await using var command = _connection!.CreateCommand();
                command.CommandText =
                    "DELETE FROM conversations WHERE conversation_id = @conversationId";
                command.Parameters.Add(
                    new SqlParameter("@conversationId", SqlDbType.NVarChar) { Value = conversationId }
                );
                await command.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error clearing conversation");
                throw;
            }
        }

        public async Task AddMessagesAsync(IEnumerable<MessageRecord> messages)
        {
            EnsureConnection();

            await using var transaction = await _connection!.BeginTransactionAsync();

            try
            {
                foreach (var message in messages)
                {
                    await using var command = _connection.CreateCommand();
                    command.Transaction = transaction as SqlTransaction;
                    command.CommandText =
                        @"INSERT INTO conversations (conversation_id, role, name, content, activity_id, timestamp, blob)
VALUES (@conversationId, @role, @name, @content, @activityId, @timestamp, @blob)";

                    command.Parameters.Add(
                        new SqlParameter("@conversationId", SqlDbType.NVarChar)
                        {
                            Value = message.ConversationId,
                        }
                    );
                    command.Parameters.Add(
                        new SqlParameter("@role", SqlDbType.NVarChar) { Value = message.Role }
                    );
                    command.Parameters.Add(
                        new SqlParameter("@name", SqlDbType.NVarChar) { Value = message.Name }
                    );
                    command.Parameters.Add(
                        new SqlParameter("@content", SqlDbType.NVarChar) { Value = message.Content }
                    );
                    command.Parameters.Add(
                        new SqlParameter("@activityId", SqlDbType.NVarChar)
                        {
                            Value = message.ActivityId,
                        }
                    );
                    command.Parameters.Add(
                        new SqlParameter("@timestamp", SqlDbType.NVarChar)
                        {
                            Value = message.Timestamp,
                        }
                    );
                    command.Parameters.Add(
                        new SqlParameter("@blob", SqlDbType.NVarChar)
                        {
                            Value = JsonSerializer.Serialize(message),
                        }
                    );

                    await command.ExecuteNonQueryAsync();
                }

                await transaction.CommitAsync();
            }
            catch
            {
                await transaction.RollbackAsync();
                throw;
            }
        }

        public async Task<int> CountMessagesAsync(string conversationId)
        {
            EnsureConnection();

            try
            {
                await using var command = _connection!.CreateCommand();
                command.CommandText =
                    "SELECT COUNT(*) FROM conversations WHERE conversation_id = @conversationId";
                command.Parameters.Add(
                    new SqlParameter("@conversationId", SqlDbType.NVarChar) { Value = conversationId }
                );

                var result = await command.ExecuteScalarAsync();
                return result != null ? Convert.ToInt32(result) : 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error counting messages");
                return 0;
            }
        }

        public Task ClearAllMessagesAsync() => ClearAllAsync();

        public async Task<List<MessageRecord>> GetFilteredMessagesAsync(
            string conversationId,
            string[] keywords,
            string startTime,
            string endTime,
            string[]? participants = null,
            int? maxResults = null
        )
        {
            EnsureConnection();

            try
            {
                var messages = new List<MessageRecord>();
                var limit = maxResults ?? 5;

                await using var command = _connection!.CreateCommand();

                var whereClauses = new List<string>
                {
                    "conversation_id = @conversationId",
                    "timestamp >= @startTime",
                    "timestamp <= @endTime",
                };

                command.Parameters.Add(
                    new SqlParameter("@conversationId", SqlDbType.NVarChar) { Value = conversationId }
                );
                command.Parameters.Add(
                    new SqlParameter("@startTime", SqlDbType.NVarChar) { Value = startTime }
                );
                command.Parameters.Add(
                    new SqlParameter("@endTime", SqlDbType.NVarChar) { Value = endTime }
                );

                if (keywords.Length > 0)
                {
                    var keywordConditions = new List<string>();
                    for (int i = 0; i < keywords.Length; i++)
                    {
                        var paramName = $"@keyword{i}";
                        keywordConditions.Add($"content LIKE {paramName}");
                        command.Parameters.Add(
                            new SqlParameter(paramName, SqlDbType.NVarChar)
                            {
                                Value = $"%{keywords[i].ToLower()}%",
                            }
                        );
                    }

                    whereClauses.Add($"({string.Join(" OR ", keywordConditions)})");
                }

                if (participants != null && participants.Length > 0)
                {
                    var participantConditions = new List<string>();
                    for (int i = 0; i < participants.Length; i++)
                    {
                        var paramName = $"@participant{i}";
                        participantConditions.Add($"name LIKE {paramName}");
                        command.Parameters.Add(
                            new SqlParameter(paramName, SqlDbType.NVarChar)
                            {
                                Value = $"%{participants[i].ToLower()}%",
                            }
                        );
                    }

                    whereClauses.Add($"({string.Join(" OR ", participantConditions)})");
                }

                command.CommandText =
                    $"SELECT TOP ({limit}) blob FROM conversations WHERE {string.Join(" AND ", whereClauses)} ORDER BY timestamp DESC";

                await using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var blob = reader.GetString(0);
                    var message = JsonSerializer.Deserialize<MessageRecord>(blob);
                    if (message != null)
                    {
                        messages.Add(message);
                    }
                }

                return messages;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving filtered messages");
                return new List<MessageRecord>();
            }
        }

        public async Task<bool> RecordFeedbackAsync(
            string replyToId,
            string reaction,
            object? feedbackJson = null
        )
        {
            EnsureConnection();

            try
            {
                await using var command = _connection!.CreateCommand();
                command.CommandText =
                    @"INSERT INTO feedback (reply_to_id, reaction, feedback)
VALUES (@replyToId, @reaction, @feedback)";

                command.Parameters.Add(
                    new SqlParameter("@replyToId", SqlDbType.NVarChar) { Value = replyToId }
                );
                command.Parameters.Add(
                    new SqlParameter("@reaction", SqlDbType.NVarChar) { Value = reaction }
                );
                command.Parameters.Add(
                    new SqlParameter("@feedback", SqlDbType.NVarChar)
                    {
                        Value = feedbackJson != null
                            ? JsonSerializer.Serialize(feedbackJson)
                            : (object)DBNull.Value,
                    }
                );

                var result = await command.ExecuteNonQueryAsync();
                return result > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error recording feedback");
                return false;
            }
        }

        public async Task CloseAsync()
        {
            if (_connection != null)
            {
                await _connection.CloseAsync();
                await _connection.DisposeAsync();
                _connection = null;
                _isInitialized = false;

                _logger.LogDebug("Closed MSSQL database connection");
            }
        }

        private void EnsureConnection()
        {
            if (_connection == null)
            {
                throw new InvalidOperationException("Database not connected");
            }
        }
    }
}
