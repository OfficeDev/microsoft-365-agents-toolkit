using Json.Schema;

namespace {{SafeProjectName}}.Capability.Template
{
    /// <summary>
    /// Defines JSON schema for Template capability functions
    ///
    /// TEMPLATE INSTRUCTIONS:
    /// Define the schema for your capability's functions here.
    /// Each schema describes the parameters the AI model must provide.
    /// </summary>
    public static class TemplateSchema
    {
        /// <summary>
        /// Arguments for the my_function_with_params function
        /// </summary>
        public class MyFunctionArgs
        {
            /// <summary>
            /// TODO: Describe this parameter
            /// </summary>
            public string Param1 { get; set; } = string.Empty;
    
            /// <summary>
            /// TODO: Describe this parameter
            /// </summary>
            public string[] Param2 { get; set; } = Array.Empty<string>();
    
            /// <summary>
            /// Maximum number of results to return
            /// </summary>
            public int Param3 { get; set; } = 10;
        }
    
        /// <summary>
        /// Example schema for a function that retrieves messages with filters
        /// TODO: Replace this with schemas for your actual functions
        /// </summary>
        public static JsonSchema GetMyFunctionSchema()
        {
            return new JsonSchemaBuilder()
                .Type(SchemaValueType.Object)
                .Properties(
                    (
                        "param1",
                        new JsonSchemaBuilder()
                            .Type(SchemaValueType.String)
                            .Description("TODO: Describe this parameter")
                            .Build()
                    ),
                    (
                        "param2",
                        new JsonSchemaBuilder()
                            .Type(SchemaValueType.Array)
                            .Items(new JsonSchemaBuilder().Type(SchemaValueType.String))
                            .Description("TODO: Describe this parameter")
                            .Build()
                    ),
                    (
                        "param3",
                        new JsonSchemaBuilder()
                            .Type(SchemaValueType.Integer)
                            .Description("Maximum number of results to return")
                            .Minimum(1)
                            .Maximum(50)
                            .Build()
                    )
                )
                .Required("param1", "param2")
                .AdditionalProperties(false)
                .Build();
        }
    }
}
