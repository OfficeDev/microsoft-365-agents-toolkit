using Json.Schema;

namespace {{SafeProjectName}}.Capability.Search
{
    /// <summary>
    /// Function schemas and parameter definitions for the search capability
    /// </summary>
    public static class SearchSchema
    {
        /// <summary>
        /// Arguments for the search_messages function
        /// </summary>
        public class SearchMessagesArgs
        {
            /// <summary>
            /// Keywords to search for in the message content
            /// </summary>
            public string[] Keywords { get; set; } = Array.Empty<string>();

            /// <summary>
            /// Optional: list of participant names to filter messages by who said them
            /// </summary>
            public string[]? Participants { get; set; }

            /// <summary>
            /// Optional: maximum number of results to return (default is 5)
            /// </summary>
            public int MaxResults { get; set; } = 5;
        }

        /// <summary>
        /// Get the JSON schema for search_messages function
        /// This defines the parameters that the AI model can use when calling the function
        /// </summary>
        public static JsonSchema GetSearchMessagesSchema()
        {
            return new JsonSchemaBuilder()
                .Type(SchemaValueType.Object)
                .Properties(
                    (
                        "keywords",
                        new JsonSchemaBuilder()
                            .Type(SchemaValueType.Array)
                            .Items(new JsonSchemaBuilder().Type(SchemaValueType.String))
                            .Description("Keywords to search for in the message content")
                            .Build()
                    ),
                    (
                        "participants",
                        new JsonSchemaBuilder()
                            .Type(SchemaValueType.Array)
                            .Items(new JsonSchemaBuilder().Type(SchemaValueType.String))
                            .Description(
                                "Optional: list of participant names to filter messages by who said them"
                            )
                            .Build()
                    ),
                    (
                        "max_results",
                        new JsonSchemaBuilder()
                            .Type(SchemaValueType.Integer)
                            .Description(
                                "Optional: maximum number of results to return (default is 5)"
                            )
                            .Minimum(1)
                            .Maximum(50)
                            .Build()
                    )
                )
                .Required("keywords")
                .AdditionalProperties(false)
                .Build();
        }
    }
}
