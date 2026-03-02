namespace {{SafeProjectName}}.Notification
{
    using Microsoft.Agents.Core.Models;
    using System.Text.Json;
    internal class LocalFileStorage: IConversationReferenceStorage
    {
        private readonly string _filePath;

        public LocalFileStorage(string dirName)
        {
            var localFileName = Environment.GetEnvironmentVariable("TEAMSFX_NOTIFICATION_STORE_FILENAME") ?? ".notification.localstore.json";
            _filePath = Path.Combine(dirName, localFileName);
        }

        public async Task<IDictionary<string, object>> ReadAsync(string[] keys, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(_filePath))
            {
                return null;
            }

            var result = new Dictionary<string, object>();
            var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
            foreach (var key in keys)
            {
                if (allData.ContainsKey(key))
                {
                    result[key] = allData[key];
                }
            }
            return result;
        }

        public async Task<IDictionary<string, TStoreItem>> ReadAsync<TStoreItem>(string[] keys, CancellationToken cancellationToken = default) where TStoreItem : class
        {
            var storeItems = await ReadAsync(keys, cancellationToken).ConfigureAwait(false);
            var values = new Dictionary<string, TStoreItem>(keys.Length);
            foreach (var entry in storeItems)
            {
                if (entry.Value is TStoreItem valueAsType)
                {
                    values.Add(entry.Key, valueAsType);
                }
            }
            return values;
        }

        public async Task<PagedData<ConversationReference>> ListAsync(int? pageSize = null, string continuationToken = null, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(_filePath))
            {
                return new PagedData<ConversationReference>
                {
                    Data = Array.Empty<ConversationReference>(),
                    ContinuationToken = null
                };
            }

            var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
            var allValues = allData.Values.ToList();

            int skip = 0;
            if (!string.IsNullOrEmpty(continuationToken) && int.TryParse(continuationToken, out int tokenValue))
            {
                skip = tokenValue;
            }

            int take = pageSize ?? allValues.Count;
            var page = allValues.Skip(skip).Take(take).ToArray();

            string nextToken = (skip + take) < allValues.Count ? (skip + take).ToString() : null;

            return new PagedData<ConversationReference>
            {
                Data = page.Cast<ConversationReference>().ToArray(),
                ContinuationToken = nextToken
            };
        }

        public async Task WriteAsync(IDictionary<string, object> changes, CancellationToken cancellationToken = default)
        {
            if (!File.Exists(_filePath))
            {
                var allData = new Dictionary<string, ConversationReference>(changes.Count);
                foreach (var kvp in changes)
                {
                    if (!(kvp.Value is ConversationReference))
                    {
                        throw new ArgumentException($"Value for key '{kvp.Key}' is not of type 'ConversationReference'.");
                    }
                    allData[kvp.Key] = (ConversationReference)kvp.Value;
                }
                await WriteToFile(allData, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
                foreach (var kvp in changes)
                {
                    allData[kvp.Key] = (ConversationReference)kvp.Value;
                }
                await WriteToFile(allData, cancellationToken).ConfigureAwait(false);
            }
        }

        public Task WriteAsync<TStoreItem>(IDictionary<string, TStoreItem> changes, CancellationToken cancellationToken = default) where TStoreItem : class
        {
            Dictionary<string, object> changesAsObject = new(changes.Count);
            foreach (var change in changes)
            {
                changesAsObject.Add(change.Key, change.Value);
            }
            return WriteAsync(changesAsObject, cancellationToken);
        }

        public async Task DeleteAsync(string[] keys, CancellationToken cancellationToken = default)
        {
            if (File.Exists(_filePath))
            {
                var allData = await ReadFromFile(cancellationToken).ConfigureAwait(false);
                foreach (var key in keys)
                {
                    if (allData.ContainsKey(key))
                    {
                        allData.Remove(key);
                    }
                }
                await WriteToFile(allData, cancellationToken).ConfigureAwait(false);
            }
        }

        private async Task<IDictionary<string, ConversationReference>> ReadFromFile(CancellationToken cancellationToken = default)
        {
            var fileInfo = new FileInfo(_filePath);
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                // return empty map
                return new Dictionary<string, ConversationReference>();
            }

            using var file = File.OpenRead(_filePath);
            return await JsonSerializer.DeserializeAsync<IDictionary<string, ConversationReference>>(file, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        private async Task WriteToFile(IDictionary<string, ConversationReference> data, CancellationToken cancellationToken = default)
        {
            using var file = File.Create(_filePath);
            await JsonSerializer.SerializeAsync(file, data, new JsonSerializerOptions { WriteIndented = true }, cancellationToken).ConfigureAwait(false);
        }
    }
}
