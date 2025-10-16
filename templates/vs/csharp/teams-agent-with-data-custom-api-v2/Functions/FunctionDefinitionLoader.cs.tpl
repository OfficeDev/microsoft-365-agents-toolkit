using System.Text.Json.Nodes;
using System.Text.Json;

namespace {{SafeProjectName}}.Functions
{
    public static class FunctionDefinitionLoader
    {

        public static readonly JsonObject FunctionDefinitions = new JsonObject();

        static FunctionDefinitionLoader()
        {
            string json = File.ReadAllText(Path.Combine("Functions","./functions.json"));
            FunctionDefinitions = JsonSerializer.Deserialize<JsonObject>(json);
        }
    }
}
