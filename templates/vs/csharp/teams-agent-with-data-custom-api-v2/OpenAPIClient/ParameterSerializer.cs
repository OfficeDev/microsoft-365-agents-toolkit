using Microsoft.OpenApi.Models;
using System.Text.Json;

namespace OpenAPIClient
{
    // serialize parameters based on schema style and explode property: https://swagger.io/specification/v3
    internal class ParameterSerializer
    {
        internal static string Serialize(JsonElement value, ParameterStyle style, bool explode, string parentKey = "")
        {
            try
            {
                switch (value.ValueKind)
                {
                    case JsonValueKind.Array:
                        return SerializeArray(value, style, explode, parentKey);
                    case JsonValueKind.Object:
                        return SerializeObject(value, style, explode, parentKey);
                    default:
                        return value.ToString();
                }
            }
            catch (Exception ex) {
                throw new SerializeParameterFailedException($"Serialize {value} with explode: {explode}, style: {style} failed due to error: " + ex.Message);
            }
        }

        private static string SerializeArray(JsonElement arrayElement, ParameterStyle style, bool explode, string parentKey)
        {
            var values = arrayElement.EnumerateArray().Select(e => Serialize(e, style, explode, parentKey)).ToList();

            if (style == ParameterStyle.Simple)
            {
                return string.Join(",", values);
            }
            else if (style == ParameterStyle.Form)
            {
                return explode ? string.Join("&", values.Select(v => $"{parentKey}={v}")) : string.Join(",", values);
            }
            else if (style == ParameterStyle.Matrix)
            {
                return explode ? string.Join(";", values.Select(v => $"{parentKey}={v}")) : string.Join(";", values);
            }
            else if (style == ParameterStyle.Label)
            {
                return explode ? string.Join(".", values.Select(v => $"{parentKey}={v}")) : string.Join(".", values);
            }
            else if (style == ParameterStyle.SpaceDelimited)
            {
                return string.Join(" ", values);
            }
            else if(style == ParameterStyle.PipeDelimited)
            {
                return string.Join("|", values);
            }

            return string.Join(",", values); // Default to simple style
        }

        private static string SerializeObject(JsonElement objectElement, ParameterStyle style, bool explode, string parentKey)
        {
            var keyValuePairs = objectElement.EnumerateObject().Select(p =>
            {
                var key = p.Name;
                var value = Serialize(p.Value, style, explode, key);
                return style == ParameterStyle.DeepObject ? $"{parentKey}[{key}]={value}" : $"{key}={value}";
            });

            if (style == ParameterStyle.Simple)
            {
                return string.Join(",", keyValuePairs);
            }
            else if (style == ParameterStyle.Form)
            {
                return explode ? string.Join("&", keyValuePairs) : string.Join(",", keyValuePairs);
            }
            else if (style == ParameterStyle.Matrix)
            {
                return explode ? string.Join(";", keyValuePairs) : string.Join(";", keyValuePairs);
            }
            else if (style == ParameterStyle.DeepObject)
            {
                return string.Join("&", keyValuePairs);
            }

            return string.Join(",", keyValuePairs); // Default to simple style
        }
    }
}
