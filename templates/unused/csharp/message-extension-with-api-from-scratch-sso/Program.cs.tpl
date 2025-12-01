using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using {{SafeProjectName}};

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureAppConfiguration((hostContext, builder) =>
    {
        var context = hostContext.HostingEnvironment;
        var configuration = new ConfigurationBuilder()
            .AddJsonFile(Path.Combine(context.ContentRootPath, "appsettings.json"), optional: true, reloadOnChange: false)
            .AddJsonFile(Path.Combine(context.ContentRootPath, $"appsettings.{context.EnvironmentName}.json"), optional: true, reloadOnChange: false)
            .Build();
        builder.AddConfiguration(configuration);
        var config = builder.Build().Get<ConfigOptions>();
        builder.AddInMemoryCollection(new Dictionary<string, string>()
        {
            { "tenantId", config.TENANT_ID },
            { "clientId", config.CLIENT_ID }
        });
    })
    .ConfigureServices((hostContext, services) =>
    {
        var configuration = hostContext.Configuration;

        services.AddSingleton(hostContext.HostingEnvironment.ContentRootPath);
        services.AddSingleton<TokenValidator>(provider =>
        {
            return new TokenValidator(configuration["tenantId"], CloudType.Public);
        });
        services.AddSingleton<AuthMiddleware>(provider =>
        {
            var tokenValidator = provider.GetRequiredService<TokenValidator>();
            var audience = configuration["clientId"];
            var issuer = $"https://login.microsoftonline.com/{configuration["tenantId"]}/v2.0";
            var allowedTenants = new[] { configuration["tenantId"] };
            return new AuthMiddleware(tokenValidator, audience, issuer, allowedTenants, null);
        });
    })
    .Build();

host.Run();