using RestSharp;
using System.Text.Json;
using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;
using RestSharp.Authenticators;

namespace OpenAPIClient
{
    public class APIClient
    {
        private RestClient RestClient;
        private OpenApiDocument Doc;

        public APIClient(string specPath)
        {
            try
            {
                using (var stream = new FileStream(specPath, FileMode.Open, FileAccess.Read))
                {
                    Doc = new OpenApiStreamReader().Read(stream, out var diagnostic);
                }
            }
            catch (Exception ex) {
                throw new ParseOpenAPISpecException("Parse OpenAPI spec file failed with error: " + ex.Message);
            }

            IAuthenticator authenticator = null;

            // You can add auth using below code
            /*
            authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(
                "YOUR_ACCESS_TOKEN", "Bearer"
            );

            authenticator = new HttpBasicAuthenticator("username", "password");

            authenticator = new JwtAuthenticator("YOUR_JWT_TOKEN");
            */

            var options = new RestClientOptions()
            {
                Authenticator = authenticator
            };

            RestClient = new RestClient(options);
        }

        public async Task<RestResponse> CallAsync(string path, Method httpMethod, RequestParams param)
        {
            OperationType operationType = MethodToOperationTypeMap[httpMethod];

            if (!Doc.Paths.ContainsKey(path) || !Doc.Paths[path].Operations.ContainsKey(operationType))
            {
                throw new APINotExistException(path, httpMethod);
            }

            var operationObj = Doc.Paths[path].Operations[operationType];
            var serverUrl = GetAPIServerUrl(path, operationType);

            if (string.IsNullOrEmpty(serverUrl) || !Uri.TryCreate(serverUrl, UriKind.Absolute, out var uriResult) ||
                (uriResult.Scheme != Uri.UriSchemeHttp && uriResult.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidServerUrlExcpetion(serverUrl);
            }

            var request = new RestRequest(serverUrl + path);

            ProcessParameters(param.PathObject, operationObj, ParameterStyle.Simple, false, (key, value) => request.AddUrlSegment(key, value));
            
            ProcessParameters(param.QueryObject, operationObj, ParameterStyle.Form, true, (key, value) => request.AddQueryParameter(key, value));
            
            ProcessParameters(param.HeaderObject, operationObj, ParameterStyle.Simple, false, (key, value) => request.AddHeader(key, value));

            if (param.RequestBody != null)
            {
                request.AddJsonBody(param.RequestBody, ContentType.Json);
            }

            var response = await RestClient.ExecuteAsync(request, httpMethod, CancellationToken.None);

            if (response.ResponseStatus == ResponseStatus.Completed && response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                return response;
            }

            throw new RequestFailedException(response);
        }

        private KeyValuePair<string, string> GetParameterKeyValuePair(JsonProperty property, OpenApiOperation operationObj , ParameterStyle defaultStyle, bool defaultExplode)
        {
            var key = property.Name;
            var value = property.Value;

            var parameterDefinition = operationObj.Parameters.FirstOrDefault(p => p.Name == key);

            var style = parameterDefinition?.Style ?? defaultStyle;
            var explode = parameterDefinition?.Explode ?? defaultExplode;

            var valueResult = ParameterSerializer.Serialize(value, style, explode, key);
            return new KeyValuePair<string, string>(key, valueResult);
        }
        
        private static readonly Dictionary<Method, OperationType> MethodToOperationTypeMap = new Dictionary<Method, OperationType>
        {
            { Method.Get, OperationType.Get },
            { Method.Post, OperationType.Post },
            { Method.Put, OperationType.Put },
            { Method.Delete, OperationType.Delete },
            { Method.Head, OperationType.Head },
            { Method.Options, OperationType.Options },
            { Method.Patch, OperationType.Patch }
        };

        private string GetAPIServerUrl(string path, OperationType operationType)
        {
            var rootServerUrl = Doc.Servers?.FirstOrDefault()?.Url;
            var apiLevelServerUrl = Doc.Paths[path].Servers?.FirstOrDefault()?.Url;
            var methodServerUrl = Doc.Paths[path].Operations[operationType].Servers?.FirstOrDefault()?.Url;
            var serverUrl = methodServerUrl ?? apiLevelServerUrl ?? rootServerUrl;
            return serverUrl;
        }

        private void ProcessParameters(object paramObj, OpenApiOperation operationObj, ParameterStyle style, bool flag, Action<string, string> addParameter)
        {
            if (paramObj != null)
            {
                var jsonElement = (JsonElement)paramObj;

                if (jsonElement.ValueKind == JsonValueKind.Object)
                {
                    foreach (JsonProperty property in jsonElement.EnumerateObject())
                    {
                        var paramKeyValuePair = GetParameterKeyValuePair(property, operationObj, style, flag);
                        addParameter(paramKeyValuePair.Key, paramKeyValuePair.Value);
                    }
                }
                else
                {
                    throw new ParameterNotObjectException(paramObj?.ToString());
                }
            }
        }
    }
}
