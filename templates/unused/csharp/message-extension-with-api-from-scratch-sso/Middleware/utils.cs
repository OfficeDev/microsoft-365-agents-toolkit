using Newtonsoft.Json.Linq;

public enum CloudType
{
    Public = 0,
    Ppe = 1,
    USGovernment = 2,
    China = 3
}

public static class JwksUriProvider
{
    public static async Task<string> GetEntraJwksUriAsync(string tenant = "common", CloudType cloud = CloudType.Public)
    {
        string cloudUrl = cloud switch
        {
            CloudType.Public => "login.microsoftonline.com",
            CloudType.Ppe => "login.windows-ppe.net",
            CloudType.USGovernment => "login.microsoftonline.us",
            CloudType.China => "login.chinacloudapi.cn",
            _ => throw new ArgumentOutOfRangeException(nameof(cloud), cloud, null)
        };

        using (var httpClient = new HttpClient())
        {
            var res = await httpClient.GetStringAsync($"https://{cloudUrl}/{tenant}/.well-known/openid-configuration");
            var data = JObject.Parse(res);
            return data["jwks_uri"]?.ToString();
        }
    }
}