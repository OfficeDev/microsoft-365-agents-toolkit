using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.AI.DataSources;
using Microsoft.Teams.AI.AI.Prompts.Sections;
using Microsoft.Teams.AI.AI.Tokenizers;
using Microsoft.Teams.AI.State;

namespace {{SafeProjectName}}
{
    public class MyDataSource : IDataSource
    {
        public string Name { get; }
        private List<string> _data = new List<string>();
        public MyDataSource(string name)
        {
            Name = name;
            Init();
        }
        public async Task<RenderedPromptSection<string>> RenderDataAsync(ITurnContext context, IMemory memory, ITokenizer tokenizer, int maxTokens, CancellationToken cancellationToken)
        {
            string? query = memory.GetValue("temp.input") as string;

            if (query == null)
            {
                return new RenderedPromptSection<string>(string.Empty, 0);
            }

            foreach (var data in _data)
            {
                if (data.Contains(query))
                {
                    //Console.WriteLine($"return rag data for data contains {query}");
                    return new RenderedPromptSection<string>(formatDocument(data), data.Length);
                }
            }
            if (query.ToLower().Contains("perksplus"))
            {
                //Console.WriteLine("return rag data for query contains perksplus");
                return new RenderedPromptSection<string>(formatDocument(_data[0]), _data[0].Length);
            }
            else if (query.ToLower().Contains("company") || query.ToLower().Contains("history"))
            {
                //Console.WriteLine("return rag data for query contains company");
                return new RenderedPromptSection<string>(formatDocument(_data[1]), _data[1].Length);
            }
            else if (query.ToLower().Contains("northwind") || query.ToLower().Contains("plan"))
            {
                //Console.WriteLine("return rag data for query contains northwind");
                return new RenderedPromptSection<string>(formatDocument(_data[2]), _data[2].Length);
            }

            return new RenderedPromptSection<string>(string.Empty, 0);
        }
        private void Init()
        {
            string[] Documents = Directory.GetFiles("data");

            int i = 0;
            foreach (string doc in Documents)
            {
                string readText = File.ReadAllText(doc);
                _data.Add(readText);
            }
            return;
        }

        private string formatDocument(string result)
        {
            return $"<context>{result}</context>";
        }
    }
}
