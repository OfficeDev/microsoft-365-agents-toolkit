using {{SafeProjectName}}.Storage.Models;
using Microsoft.Data.SqlClient;

namespace {{SafeProjectName}}.Storage;

public class SqlServerConversationStorage : IConversationStorage
{
    private readonly string _connectionString;
    private bool _initialized;

    public SqlServerConversationStorage(string connectionString)
    {
        _connectionString = connectionString;
    }

    public async Task InitializeAsync()
    {
        if (_initialized) return;

        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var createTables = @"
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Messages' AND xtype='U')
            CREATE TABLE Messages (
                id NVARCHAR(255) PRIMARY KEY,
                conversationId NVARCHAR(255) NOT NULL,
                name NVARCHAR(255) NOT NULL,
                content NVARCHAR(MAX) NOT NULL,
                timestamp DATETIME2 NOT NULL,
                role NVARCHAR(50) NOT NULL
            );

            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'idx_conversation')
            CREATE INDEX idx_conversation ON Messages(conversationId);

            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'idx_timestamp')
            CREATE INDEX idx_timestamp ON Messages(timestamp);

            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Feedback' AND xtype='U')
            CREATE TABLE Feedback (
                messageId NVARCHAR(255) PRIMARY KEY,
                reaction NVARCHAR(50) NOT NULL,
                feedbackJson NVARCHAR(MAX),
                timestamp DATETIME2 NOT NULL
            );
        ";

        using var cmd = connection.CreateCommand();
        cmd.CommandText = createTables;
        await cmd.ExecuteNonQueryAsync();

        _initialized = true;
    }

    public async Task<bool> AddMessageAsync(ConversationMessage message)
    {
        using var connection = new SqlConnection(_connectionString);
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
        cmd.Parameters.AddWithValue("@timestamp", message.Timestamp);
        cmd.Parameters.AddWithValue("@role", message.Role);

        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> AddMessagesAsync(IEnumerable<ConversationMessage> messages)
    {
        using var connection = new SqlConnection(_connectionString);
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
                cmd.Parameters.AddWithValue("@timestamp", message.Timestamp);
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
        using var connection = new SqlConnection(_connectionString);
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
            cmd.Parameters.AddWithValue("@endTime", endTime.Value);
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
        cmd.Parameters.AddWithValue("@startTime", startTime);

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
                Timestamp = reader.GetDateTime(4),
                Role = reader.GetString(5)
            });
        }

        return messages;
    }

    public async Task<IEnumerable<ConversationMessage>> GetAllMessagesAsync(string conversationId)
    {
        using var connection = new SqlConnection(_connectionString);
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
                Timestamp = reader.GetDateTime(4),
                Role = reader.GetString(5)
            });
        }

        return messages;
    }

    public async Task<bool> ClearConversationAsync(string conversationId)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        cmd.CommandText = "DELETE FROM Messages WHERE conversationId = @conversationId";
        cmd.Parameters.AddWithValue("@conversationId", conversationId);

        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public async Task<bool> RecordFeedbackAsync(string messageId, string reaction, string? feedbackJson = null)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        using var cmd = connection.CreateCommand();
        cmd.CommandText = @"
            MERGE Feedback AS target
            USING (SELECT @messageId AS messageId) AS source
            ON target.messageId = source.messageId
            WHEN MATCHED THEN
                UPDATE SET reaction = @reaction, feedbackJson = @feedbackJson, timestamp = @timestamp
            WHEN NOT MATCHED THEN
                INSERT (messageId, reaction, feedbackJson, timestamp)
                VALUES (@messageId, @reaction, @feedbackJson, @timestamp);
        ";

        cmd.Parameters.AddWithValue("@messageId", messageId);
        cmd.Parameters.AddWithValue("@reaction", reaction);
        cmd.Parameters.AddWithValue("@feedbackJson", feedbackJson ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("@timestamp", DateTime.UtcNow);

        return await cmd.ExecuteNonQueryAsync() > 0;
    }
}
