using {{SafeProjectName}};
using Azure;
using Azure.Search.Documents;
using Azure.Search.Documents.Models;
using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.AI.DataSources;
using Microsoft.Teams.AI.AI.Embeddings;
using Microsoft.Teams.AI.AI.Prompts.Sections;
using Microsoft.Teams.AI.AI.Tokenizers;
using Microsoft.Teams.AI.State;
using System.Text;

namespace {{SafeProjectName}}
{
    public class AzureAISearchDataSource : IDataSource
    {
        public string Name { get; }

        public readonly AzureAISearchDataSourceOptions Options;

        public readonly SearchClient SearchClient;

        public AzureAISearchDataSource(AzureAISearchDataSourceOptions options)
        {
            Options = options;
            Name = options.Name;

            AzureKeyCredential credential = new AzureKeyCredential(options.AzureAISearchApiKey);
            SearchClient = new SearchClient(options.AzureAISearchEndpoint, options.IndexName, credential);
        }

        public async Task<RenderedPromptSection<string>> RenderDataAsync(ITurnContext context, IMemory memory, ITokenizer tokenizer, int maxTokens, CancellationToken cancellationToken = default)
        {
            string query = (string)memory.GetValue("temp.input")!;

            if (string.IsNullOrEmpty(query))
            {
                return new RenderedPromptSection<string>("");
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


            // Concatenate the restaurant documents (i.e json object) string into a single document
            // until the maximum token limit is reached. This can be specified in the prompt template.
            int usedTokens = tokenizer.Encode("Contexts: ").Count;
            StringBuilder doc = new StringBuilder("Contexts: ");
            Pageable<SearchResult<Document>> results = search.GetResults();
            foreach (SearchResult<Document> result in results)
            {
                string document = $"<context>{result.Document}</context>";
                int tokens = tokenizer.Encode(document).Count;

                if (usedTokens + tokens > maxTokens)
                {
                    break;
                }

                doc.Append(document);
                usedTokens += tokens;
            }

            return new RenderedPromptSection<string>(doc.ToString(), usedTokens, usedTokens > maxTokens);
        }

        private async Task<ReadOnlyMemory<float>> _GetEmbeddingVector(string query)
        {
{{#useOpenAI}}
            OpenAIEmbeddingsOptions options = new(this.Options.OpenAIApiKey, this.Options.OpenAIEmbeddingModel);
{{/useOpenAI}}
{{#useAzureOpenAI}}
            AzureOpenAIEmbeddingsOptions options = new(this.Options.AzureOpenAIApiKey, this.Options.AzureOpenAIEmbeddingDeployment, this.Options.AzureOpenAIEndpoint);
{{/useAzureOpenAI}}
            OpenAIEmbeddings embeddings = new(options);
            EmbeddingsResponse response = await embeddings.CreateEmbeddingsAsync(new List<string> { query });

            return response.Output!.First();
        }
    }

    public class AzureAISearchDataSourceOptions
    {
        /// <summary>
        /// Name of the data source
        /// </summary>
        public string Name { get; set; }

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
