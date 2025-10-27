using {{SafeProjectName}}.Storage.Models;
using Microsoft.Data.Sqlite;

namespace {{SafeProjectName}}.Storage;

public class SqliteConversationStorage : IConversationStorage
{
    private readonly string _connectionString;
    private bool _initialized;

    public SqliteConversationStorage(string connectionString)
    {
        _connectionString = connectionString;
    }

    public async Task InitializeAsync()
    {
        if (_initialized) return;

        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        // Create messages table
        var createMessagesTable = @"
            CREATE TABLE IF NOT EXISTS Messages (
                id TEXT PRIMARY KEY,
                conversationId TEXT NOT NULL,
                name TEXT NOT NULL,
                content TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                role TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_conversation ON Messages(conversationId);
            CREATE INDEX IF NOT EXISTS idx_timestamp ON Messages(timestamp);
        ";

        // Create feedback table
        var createFeedbackTable = @"
            CREATE TABLE IF NOT EXISTS Feedback (
                messageId TEXT PRIMARY KEY,
                reaction TEXT NOT NULL,
                feedbackJson TEXT,
                timestamp TEXT NOT NULL
            );
        ";

        using var cmd = connection.CreateCommand();
        cmd.CommandText = createMessagesTable + createFeedbackTable;
        await cmd.ExecuteNonQueryAsync();

        _initialized = true;
    }

    public async Task<bool> AddMessageAsync(ConversationMessage message)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO Messages (id, conversationId, name, content, timestamp, role)
            VALUES (@id, @conversationId, @name, @content, @timestamp, @role)
        ";

        cmd.Parameters.AddWithValue("@id", message.Id);
        cmd.Parameters.AddWithValue("@conversationId", message.ConversationId);
        cmd.Parameters.AddWithValue("@name", message.Name);
        cmd.Parameters.AddWithValue("@content", message.Content);
        cmd.Parameters.AddWithValue("@timestamp", message.Timestamp.ToString("o"));
        cmd.Parameters.AddWithValue("@role", message.Role);

        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> AddMessagesAsync(IEnumerable<ConversationMessage> messages)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var transaction = connection.BeginTransaction();
        try
        {
            foreach (var message in messages)
            {
                using var cmd = connection.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = @"
                    INSERT INTO Messages (id, conversationId, name, content, timestamp, role)
                    VALUES (@id, @conversationId, @name, @content, @timestamp, @role)
                ";

                cmd.Parameters.AddWithValue("@id", message.Id);
                cmd.Parameters.AddWithValue("@conversationId", message.ConversationId);
                cmd.Parameters.AddWithValue("@name", message.Name);
                cmd.Parameters.AddWithValue("@content", message.Content);
                cmd.Parameters.AddWithValue("@timestamp", message.Timestamp.ToString("o"));
                cmd.Parameters.AddWithValue("@role", message.Role);

                await cmd.ExecuteNonQueryAsync();
            }

            transaction.Commit();
            return true;
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    public async Task<IEnumerable<ConversationMessage>> GetMessagesByTimeRangeAsync(
        string conversationId, DateTime startTime, DateTime? endTime = null)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        
        if (endTime.HasValue)
        {
            cmd.CommandText = @"
                SELECT id, conversationId, name, content, timestamp, role
                FROM Messages
                WHERE conversationId = @conversationId 
                  AND timestamp >= @startTime 
                  AND timestamp <= @endTime
                ORDER BY timestamp ASC
            ";
            cmd.Parameters.AddWithValue("@endTime", endTime.Value.ToString("o"));
        }
        else
        {
            cmd.CommandText = @"
                SELECT id, conversationId, name, content, timestamp, role
                FROM Messages
                WHERE conversationId = @conversationId 
                  AND timestamp >= @startTime
                ORDER BY timestamp ASC
            ";
        }

        cmd.Parameters.AddWithValue("@conversationId", conversationId);
        cmd.Parameters.AddWithValue("@startTime", startTime.ToString("o"));

        var messages = new List<ConversationMessage>();
        using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            messages.Add(new ConversationMessage
            {
                Id = reader.GetString(0),
                ConversationId = reader.GetString(1),
                Name = reader.GetString(2),
                Content = reader.GetString(3),
                Timestamp = DateTime.Parse(reader.GetString(4)),
                Role = reader.GetString(5)
            });
        }

        return messages;
    }

    public async Task<IEnumerable<ConversationMessage>> GetAllMessagesAsync(string conversationId)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"
            SELECT id, conversationId, name, content, timestamp, role
            FROM Messages
            WHERE conversationId = @conversationId
            ORDER BY timestamp ASC
        ";
        cmd.Parameters.AddWithValue("@conversationId", conversationId);

        var messages = new List<ConversationMessage>();
        using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            messages.Add(new ConversationMessage
            {
                Id = reader.GetString(0),
                ConversationId = reader.GetString(1),
                Name = reader.GetString(2),
                Content = reader.GetString(3),
                Timestamp = DateTime.Parse(reader.GetString(4)),
                Role = reader.GetString(5)
            });
        }

        return messages;
    }

    public async Task<bool> ClearConversationAsync(string conversationId)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        cmd.CommandText = "DELETE FROM Messages WHERE conversationId = @conversationId";
        cmd.Parameters.AddWithValue("@conversationId", conversationId);

        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> RecordFeedbackAsync(string messageId, string reaction, string? feedbackJson = null)
    {
        using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"
            INSERT OR REPLACE INTO Feedback (messageId, reaction, feedbackJson, timestamp)
            VALUES (@messageId, @reaction, @feedbackJson, @timestamp)
        ";

        cmd.Parameters.AddWithValue("@messageId", messageId);
        cmd.Parameters.AddWithValue("@reaction", reaction);
        cmd.Parameters.AddWithValue("@feedbackJson", feedbackJson ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@timestamp", DateTime.UtcNow.ToString("o"));

        return await cmd.ExecuteNonQueryAsync() > 0;
    }
}
