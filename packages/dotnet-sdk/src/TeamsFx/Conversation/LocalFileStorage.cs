// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;
    using System.Text.Json;

    [Obsolete]
    internal sealed class LocalFileStorage : INotificationTargetStorage
    {
        private readonly string _filePath;

        public LocalFileStorage(string dirName)
        {
            var localFileName = Environment.GetEnvironmentVariable("TEAMSFX_NOTIFICATION_STORE_FILENAME") ?? ".notification.localstore.json";
            _filePath = Path.Combine(dirName, localFileName);
        }

        public async Task<ConversationReference> Read(string key, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(_filePath))
            {
                return null;
            }

            var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
            if (allData.ContainsKey(key))
            {
                return allData[key];
            }
            else
            {
                return null;
            }
        }

        public async Task<ConversationReference[]> List(CancellationToken cancellationToken = default)
        {
            if (!File.Exists(_filePath))
            {
                return Array.Empty<ConversationReference>();
            }

            var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
            return allData.Values.ToArray();
        }

        public async Task Write(string key, ConversationReference reference, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(_filePath))
            {
                await WriteToFile(new Dictionary<string, ConversationReference> { { key, reference } }, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
                allData[key] = reference;
                await WriteToFile(allData, cancellationToken).ConfigureAwait(false);
            }
        }

        public async Task Delete(string key, CancellationToken cancellationToken = default)
        {
            if (File.Exists(_filePath))
            {
                var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
                if (allData.ContainsKey(key))
                {
                    allData.Remove(key);
                    await WriteToFile(allData, cancellationToken).ConfigureAwait(false);
                }
            }
        }

        private async Task<Dictionary<string, ConversationReference>> ReadFromFile(CancellationToken cancellationToken = default)
        {
            var fileInfo = new FileInfo(_filePath);
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                // return empty map
                return new Dictionary<string, ConversationReference>();
            }

            using var file = File.OpenRead(_filePath);
            return await JsonSerializer.DeserializeAsync<Dictionary<string, ConversationReference>>(file, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        private async Task WriteToFile(Dictionary<string, ConversationReference> data, CancellationToken cancellationToken = default)
        {
            using var file = File.Create(_filePath);
            await JsonSerializer.SerializeAsync(file, data, new JsonSerializerOptions { WriteIndented = true }, cancellationToken).ConfigureAwait(false);
        }
    }
}
