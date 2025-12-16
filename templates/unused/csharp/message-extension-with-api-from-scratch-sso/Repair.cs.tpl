using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace {{SafeProjectName}}
{
    public class Repair
    {
        private readonly ILogger _logger;
        private readonly AuthMiddleware _authMiddleware;

        public Repair(ILoggerFactory loggerFactory, AuthMiddleware authMiddleware)
        {
            _logger = loggerFactory.CreateLogger<Repair>();
            _authMiddleware = authMiddleware;
        }

        [Function("repair")]
        public async Task<HttpResponseData> RunAsync([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
        {
            if (!await _authMiddleware.ValidateTokenAsync(req, _logger))
            {
                var unauthorizedResponse = req.CreateResponse(System.Net.HttpStatusCode.Unauthorized);
                return unauthorizedResponse;
            }
            
            // Log that the HTTP trigger function received a request.
            _logger.LogInformation("C# HTTP trigger function processed a request.");

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
    }
}