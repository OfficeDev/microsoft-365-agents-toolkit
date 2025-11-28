using System.Data;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace {{SafeProjectName}}.Storage
{
    public class SqliteDatabase : IDatabase
    {
        private readonly ILogger _logger;
        private readonly string _dbPath;
        private SqliteConnection? _connection;

        public SqliteDatabase(ILogger logger, string? dbPath = null)
        {
            _logger = logger;
            _dbPath =
                Environment.GetEnvironmentVariable("CONVERSATIONS_DB_PATH")
                ?? dbPath
                ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "conversations.db");

            _logger.LogDebug("SQLite database path: {DbPath}", _dbPath);
        }

        public async Task InitializeAsync()
        {
            try
            {
                var directory = Path.GetDirectoryName(_dbPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                _connection = new SqliteConnection($"Data Source={_dbPath}");
                await _connection.OpenAsync();

                await using var command = _connection.CreateCommand();

                command.CommandText =
                    @"CREATE TABLE IF NOT EXISTS conversations (
    conversation_id TEXT NOT NULL,
    role TEXT NOT NULL,
    name TEXT NOT NULL,
    content TEXT NOT NULL,
    activity_id TEXT NOT NULL,
    timestamp TEXT NOT NULL,
    blob TEXT NOT NULL
)";
                await command.ExecuteNonQueryAsync();

                command.CommandText =
                    @"CREATE INDEX IF NOT EXISTS idx_conversation_id ON conversations(conversation_id)";
                await command.ExecuteNonQueryAsync();

                command.CommandText =
                    @"CREATE TABLE IF NOT EXISTS feedback (
    reply_to_id TEXT NOT NULL,
    reaction TEXT NOT NULL CHECK (reaction IN ('like','dislike')),
    feedback TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
)";
                await command.ExecuteNonQueryAsync();

                _logger.LogDebug("SQLite database initialized");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error initializing SQLite database");
                throw;
            }
        }

        public async Task ClearAllAsync()
        {
            EnsureConnection();

            await using var command = _connection!.CreateCommand();
            command.CommandText = "DELETE FROM conversations; VACUUM;";
            await command.ExecuteNonQueryAsync();

            _logger.LogDebug("Cleared all conversations from SQLite store");
        }

        public async Task<List<MessageRecord>> GetAsync(string conversationId)
        {
            EnsureConnection();

            var messages = new List<MessageRecord>();

            await using var command = _connection!.CreateCommand();
            command.CommandText =
                "SELECT blob FROM conversations WHERE conversation_id = @conversationId ORDER BY timestamp ASC";
            command.Parameters.AddWithValue("@conversationId", conversationId);

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

        public async Task<List<MessageRecord>> GetMessagesByTimeRangeAsync(
            string conversationId,
            string startTime,
            string endTime
        )
        {
            EnsureConnection();

            var messages = new List<MessageRecord>();

            await using var command = _connection!.CreateCommand();
            command.CommandText =
                @"SELECT blob FROM conversations
WHERE conversation_id = @conversationId
    AND timestamp >= @startTime
    AND timestamp <= @endTime
ORDER BY timestamp ASC";
            command.Parameters.AddWithValue("@conversationId", conversationId);
            command.Parameters.AddWithValue("@startTime", startTime);
            command.Parameters.AddWithValue("@endTime", endTime);

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

            await using var command = _connection!.CreateCommand();
            command.CommandText = "DELETE FROM conversations WHERE conversation_id = @conversationId";
            command.Parameters.AddWithValue("@conversationId", conversationId);
            await command.ExecuteNonQueryAsync();
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
                    command.Transaction = transaction as SqliteTransaction;
                    command.CommandText =
                        @"INSERT INTO conversations (conversation_id, role, name, content, activity_id, timestamp, blob)
VALUES (@conversationId, @role, @name, @content, @activityId, @timestamp, @blob)";

                    command.Parameters.AddWithValue("@conversationId", message.ConversationId);
                    command.Parameters.AddWithValue("@role", message.Role);
                    command.Parameters.AddWithValue("@name", message.Name);
                    command.Parameters.AddWithValue("@content", message.Content);
                    command.Parameters.AddWithValue("@activityId", message.ActivityId);
                    command.Parameters.AddWithValue("@timestamp", message.Timestamp);
                    command.Parameters.AddWithValue("@blob", JsonSerializer.Serialize(message));

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

            await using var command = _connection!.CreateCommand();
            command.CommandText =
                "SELECT COUNT(*) FROM conversations WHERE conversation_id = @conversationId";
            command.Parameters.AddWithValue("@conversationId", conversationId);

            var result = await command.ExecuteScalarAsync();
            return result != null ? Convert.ToInt32(result) : 0;
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

            var messages = new List<MessageRecord>();
            var limit = maxResults ?? 5;

            var whereClauses = new List<string>
            {
                "conversation_id = @conversationId",
                "timestamp >= @startTime",
                "timestamp <= @endTime",
            };

            if (keywords.Length > 0)
            {
                var keywordClauses = string.Join(
                    " OR ",
                    keywords.Select((_, i) => $"content LIKE @keyword{i}")
                );
                whereClauses.Add($"({keywordClauses})");
            }

            if (participants != null && participants.Length > 0)
            {
                var participantClauses = string.Join(
                    " OR ",
                    participants.Select((_, i) => $"name LIKE @participant{i}")
                );
                whereClauses.Add($"({participantClauses})");
            }

            await using var command = _connection!.CreateCommand();
            command.CommandText =
                $"SELECT blob FROM conversations WHERE {string.Join(" AND ", whereClauses)} ORDER BY timestamp DESC LIMIT @limit";

            command.Parameters.AddWithValue("@conversationId", conversationId);
            command.Parameters.AddWithValue("@startTime", startTime);
            command.Parameters.AddWithValue("@endTime", endTime);
            command.Parameters.AddWithValue("@limit", limit);

            for (int i = 0; i < keywords.Length; i++)
            {
                command.Parameters.AddWithValue($"@keyword{i}", $"%{keywords[i].ToLower()}%");
            }

            if (participants != null)
            {
                for (int i = 0; i < participants.Length; i++)
                {
                    command.Parameters.AddWithValue($"@participant{i}", $"%{participants[i].ToLower()}%");
                }
            }

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

                command.Parameters.AddWithValue("@replyToId", replyToId);
                command.Parameters.AddWithValue("@reaction", reaction);
                command.Parameters.AddWithValue(
                    "@feedback",
                    feedbackJson != null ? JsonSerializer.Serialize(feedbackJson) : DBNull.Value
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

                _logger.LogDebug("Closed SQLite database connection");
            }
        }

        private void EnsureConnection()
        {
            if (_connection == null)
            {
                throw new InvalidOperationException("Database not initialized");
            }
        }
    }
}
