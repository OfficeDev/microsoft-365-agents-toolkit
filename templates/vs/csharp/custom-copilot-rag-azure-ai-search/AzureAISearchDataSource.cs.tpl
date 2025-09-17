using Azure;
using Azure.Search.Documents;
using Azure.Search.Documents.Models;
using System.Text;
using OpenAI.Embeddings;
using System.ClientModel;

namespace {{SafeProjectName}}
{
    public class AzureAISearchDataSource
    {
        public readonly AzureAISearchDataSourceOptions Options;

        public readonly SearchClient SearchClient;

        public AzureAISearchDataSource(AzureAISearchDataSourceOptions options)
        {
            Options = options;

            AzureKeyCredential credential = new AzureKeyCredential(options.AzureAISearchApiKey);
            SearchClient = new SearchClient(options.AzureAISearchEndpoint, options.IndexName, credential);
        }

        public async Task<string> RenderDataAsync(string query)
        {
            if (string.IsNullOrEmpty(query))
            {
                return string.Empty;
            }

            List<string> selectedFields = new() { "DocId", "DocTitle", "Description" };
            List<string> searchFields = new() { "DocTitle", "Description" };

            //// HYBRID SEARCH ////
            //// Search using both vector and text search
            SearchOptions options = new();
            ReadOnlyMemory<float> vectorizedQuery = await this._GetEmbeddingVector(query);
            foreach (string field in searchFields)
            {
                options.SearchFields.Add(field);
            }

            foreach (string field in selectedFields)
            {
                options.Select.Add(field);
            }
            options.VectorSearch = new()
            {
                Queries = { new VectorizedQuery(vectorizedQuery) { KNearestNeighborsCount = 3, Fields = { "DescriptionVector" } } }
            };
            SearchResults<Document> search = SearchClient.Search<Document>(query, options);

            StringBuilder doc = new StringBuilder("Contexts: ");
            Pageable<SearchResult<Document>> results = search.GetResults();
            foreach (SearchResult<Document> result in results)
            {
                string document = $"<context>{result.Document}</context>";

                doc.Append(document);
            }

            return doc.ToString();
        }

        private async Task<ReadOnlyMemory<float>> _GetEmbeddingVector(string query)
        {
{{#useOpenAI}}
            EmbeddingClient client = new(this.Options.OpenAIEmbeddingModel, this.Options.OpenAIApiKey);
{{/useOpenAI}}
{{#useAzureOpenAI}}
            EmbeddingClient client = new(this.Options.AzureOpenAIEmbeddingDeployment, new ApiKeyCredential(this.Options.AzureOpenAIApiKey), new OpenAI.OpenAIClientOptions()
            {
                Endpoint = new Uri($"{this.Options.AzureOpenAIEndpoint}/openai/v1")
            });
{{/useAzureOpenAI}}
           
            OpenAIEmbedding embedding = await client.GenerateEmbeddingAsync(query);
            return embedding.ToFloats();
        }
    }

    public class AzureAISearchDataSourceOptions
    {
        /// <summary>
        /// Name of the Azure AI Search index
        /// </summary>
        public string IndexName { get; set; }

        /// <summary>
        /// Azure AI Search API key
        /// </summary>
        public string AzureAISearchApiKey { get; set; }

        /// <summary>
        /// Azure AI Search endpoint
        /// </summary>
        public Uri AzureAISearchEndpoint { get; set; }
        
{{#useOpenAI}}
        /// <summary>
        /// OpenAI API key
        /// </summary>
        public string OpenAIApiKey { get; set; }

        /// <summary>
        /// OpenAI embeddings deployment name
        /// </summary>
        public string OpenAIEmbeddingModel { get; set; }
{{/useOpenAI}}
{{#useAzureOpenAI}}
        /// <summary>
        /// Azure OpenAI API key
        /// </summary>
        public string AzureOpenAIApiKey { get; set; }

        /// <summary>
        /// Azure OpenAI endpoint
        /// </summary>
        public string AzureOpenAIEndpoint { get; set; }

        /// <summary>
        /// Azure OpenAI deployment name
        /// </summary>
        public string AzureOpenAIDeploymentName { get; set; }

        /// <summary>
        /// Azure OpenAI embeddings deployment name
        /// </summary>
        public string AzureOpenAIEmbeddingDeployment { get; set; }
{{/useAzureOpenAI}}
    }
}
