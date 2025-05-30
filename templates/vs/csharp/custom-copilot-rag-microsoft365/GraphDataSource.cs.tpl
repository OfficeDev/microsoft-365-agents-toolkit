using Azure;
using Microsoft.Bot.Builder;
using Microsoft.Graph.SecurityNamespace;
using Microsoft.Teams.AI.AI.DataSources;
using Microsoft.Teams.AI.AI.Prompts.Sections;
using Microsoft.Teams.AI.AI.Tokenizers;
using Microsoft.Teams.AI.State;
using System.Reflection.Metadata;
using System.Text;

namespace {{SafeProjectName}}
{
    public class GraphDataSource : IDataSource
    {
        public string Name { get; }
        public GraphDataSource(string name)
        {
            Name = name;
        }
        public async Task<RenderedPromptSection<string>> RenderDataAsync(ITurnContext context, IMemory memory, ITokenizer tokenizer, int maxTokens, CancellationToken cancellationToken)
        {
            string? query = memory.GetValue("temp.input") as string;

            if (query == null)
            {
                return new RenderedPromptSection<string>(string.Empty, 0);
            }
            var token = (memory as TurnState).Temp.AuthTokens["graph"];
            var graphClient = new SimpleGraphClient(token);
           
            var graphQuery = query;
            if (query.ToLower().Contains("perksplus"))
            {
                graphQuery = "perksplus program";
            }
            else if (query.ToLower().Contains("company"))
            {
                graphQuery = "company history";
            }
            else if (query.ToLower().Contains("northwind"))
            {
                graphQuery = "northwind health";
            }

            var resourceUrls = await graphClient.GetQuery(graphQuery);
            // Concatenate the restaurant documents (i.e json object) string into a single document
            // until the maximum token limit is reached. This can be specified in the prompt template.
            int usedTokens = 0;
            StringBuilder doc = new StringBuilder("");
            foreach (var url in resourceUrls)
            {
                var content = await graphClient.DownloadSharepointFile(SimpleGraphClient.UrlToSharingToken(url));
                string document = $"{content}\n\n";
                int tokens = tokenizer.Encode(document).Count;

                if (usedTokens + tokens > maxTokens)
                {
                    break;
                }

                doc.Append(document);
                usedTokens += tokens;
            }

            return new RenderedPromptSection<string>(formatDocument(doc.ToString()), usedTokens, usedTokens > maxTokens);
        }

        private string formatDocument(string result)
        {
            return $"<context>{result}</context>";
        }
    }
}
