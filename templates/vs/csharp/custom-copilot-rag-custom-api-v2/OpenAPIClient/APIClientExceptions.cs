using RestSharp;

namespace OpenAPIClient
{
    public class APIClientException: Exception
    {
        public APIClientException(string message) : base(message)
        {
        }
    }

    public class ParseOpenAPISpecException : APIClientException
    {
        public ParseOpenAPISpecException(string message) : base(message)
        {
        }
    }

    public class APINotExistException : APIClientException
    {
        public APINotExistException(string apiPath, Method httpMethod)
            : base($"API {httpMethod.ToString()} {apiPath} does not exist in the OpenAPI specification file.")
        {
        }
    }

    public class InvalidServerUrlExcpetion : APIClientException
    {
        public InvalidServerUrlExcpetion(string serverUrl)
            : base($"Server URL '{serverUrl}' is invalid. It should use the HTTP/HTTPS protocol with an absolute path.")
        {
        }
    }

    public class ParameterNotObjectException : APIClientException
    {
        public ParameterNotObjectException(string paramInfo) : base($"Parameter: {paramInfo} is not an object.")
        {
        }
    }

    public class SerializeParameterFailedException : APIClientException
    {
        public SerializeParameterFailedException(string message) : base(message)
        {
        }
    }

    public class RequestFailedException : APIClientException
    {
        public RequestFailedException(RestResponse response) : base($"Request failed with status: {response.ResponseStatus}, status code: {response.StatusCode}, error message: {response.ErrorMessage}")
        {
        }
    }
}
