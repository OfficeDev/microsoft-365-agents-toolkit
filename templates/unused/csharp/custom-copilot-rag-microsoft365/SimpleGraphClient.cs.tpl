using Microsoft.Graph;
using System.Net.Http.Headers;

namespace {{SafeProjectName}}
{
    public class SimpleGraphClient
    {
        private readonly string _token;
        private GraphServiceClient graphClient;

        public SimpleGraphClient(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new ArgumentNullException(nameof(token));
            }

            _token = token;
            graphClient = GetAuthenticatedClient();
        }

        public async Task<List<string>> GetQuery(string query)
        {
            var q = new SearchRequestObject()
            {
                EntityTypes = new[]
                {
                    EntityType.DriveItem
                },
                Query = new SearchQuery()
                {
                    QueryString = query
                },
            };
            var res = await graphClient.Search.Query(new List<SearchRequestObject> { q }).Request().PostAsync();
            var urls = res.CurrentPage[0].HitsContainers.SelectMany(hitContainer => hitContainer.Hits ?? new List<SearchHit>()).Select(hit => (hit.Resource as DriveItem).WebUrl).ToList();
            return urls;
        }

        public async Task<string> DownloadSharepointFile(string contentUrl)
        {
            var fileContentResponse = await graphClient.HttpProvider.SendAsync(new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/shares/{contentUrl}/driveItem/content")
            {
                Headers =
                {
                Authorization = new AuthenticationHeaderValue("Bearer", _token)
            }
            });

            var fileContent = await fileContentResponse.Content.ReadAsStringAsync();
            return fileContent;
        }

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        private GraphServiceClient GetAuthenticatedClient()
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        return Task.CompletedTask;
                    }));
            return graphClient;
        }

        public static string UrlToSharingToken(string inputUrl)
        {
            var base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(inputUrl));
            return "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
        }
    }
}
