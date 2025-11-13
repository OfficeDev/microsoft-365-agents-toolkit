using System.ComponentModel;
using System.Text.Json.Serialization;

namespace TravelAgent.Bot.Agents;

public enum TravelAgentResponseContentType
{
    [JsonPropertyName("text")]
    Text,

    [JsonPropertyName("adaptive-card")]
    AdaptiveCard
}

public class TravelAgentResponse
{
    [JsonPropertyName("contentType")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public TravelAgentResponseContentType ContentType { get; set; }

    [JsonPropertyName("content")]
    [Description("The content of the response in plain text,or JSON based adaptive card but must be a string.")]
    public string Content { get; set; }
}
