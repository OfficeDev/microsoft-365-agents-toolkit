using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Net;

namespace {{SafeProjectName}}
{
    public class Repair
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;

        public Repair(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _logger = loggerFactory.CreateLogger<Repair>();
            _configuration = configuration;
        }

        [Function("repair")]
        public async Task<HttpResponseData> RunAsync([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
        {
            // Log that the HTTP trigger function received a request.
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            // Check if the API key is valid.
            if (!IsApiKeyValid(req))
            {
                // Return a 401 Unauthorized response if the API key is invalid.
                return req.CreateResponse(HttpStatusCode.Unauthorized);
            }

            // Get the query parameters from the request.
            string assignedTo = req.Query["assignedTo"];

            // Get the repair records.
            var repairRecords = RepairData.GetRepairs();

            // Filter the repair records by the assignedTo query parameter.
            var repairs = repairRecords.Where(r =>
            {
                // Split assignedTo into firstName and lastName
                var parts = r.AssignedTo.Split(' ');

                // Check if the assignedTo query parameter matches the repair record's assignedTo value, or the repair record's firstName or lastName.
                return r.AssignedTo.Equals(assignedTo?.Trim(), StringComparison.InvariantCultureIgnoreCase) ||
                       parts[0].Equals(assignedTo?.Trim(), StringComparison.InvariantCultureIgnoreCase) ||
                       parts[1].Equals(assignedTo?.Trim(), StringComparison.InvariantCultureIgnoreCase);
            });

            // Return filtered repair records, or an empty array if no records were found.
            var response = req.CreateResponse();
            await response.WriteAsJsonAsync(new { results = repairs });
            return response;
        }

        /**
        * The reason for this implementation is that Azure Function Core Tools does not support authentication when running locally.
        * This template is designed to demonstrate and facilitate local debugging of authentication functionalities in the API-based
        * message extension. Therefore, this approach was taken. If you prefer to leverage the Azure Functions' built-in API key
        * authentication, please refer to https://aka.ms/function-key-csharp for guidance.
        */
        private bool IsApiKeyValid(HttpRequestData req)
        {
            // Try to get the value of the 'Authorization' header from the request.
            // If the header is not present, return false.
            if (!req.Headers.TryGetValues("Authorization", out var authValue))
            {
                return false;
            }

            // Get the api key value from the 'Authorization' header.
            var apiKey = authValue.FirstOrDefault().Replace("Bearer", "").Trim();

            // Get the API key from the configuration.
            var configApiKey = _configuration["API_KEY"];

            // Check if the API key from the request matches the API key from the configuration.
            return apiKey == configApiKey;
        }
    }
}