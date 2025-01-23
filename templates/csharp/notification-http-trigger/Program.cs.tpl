using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.TeamsFx.Conversation;
using {{SafeProjectName}};


var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureAppConfiguration((hostContext, builder) => {
        var context = hostContext.HostingEnvironment;
        var configuration = new ConfigurationBuilder()
            .AddJsonFile(Path.Combine(context.ContentRootPath, $"appsettings.json"), optional: true, reloadOnChange: false)
            .AddJsonFile(Path.Combine(context.ContentRootPath, $"appsettings.{context.EnvironmentName}.json"), optional: true, reloadOnChange: false)
            .Build();
        builder.AddConfiguration(configuration);
        var config = builder.Build().Get<ConfigOptions>();
        builder.AddInMemoryCollection(new Dictionary<string, string>()
        {
            { "MicrosoftAppType", config.BOT_TYPE },
            { "MicrosoftAppId", config.BOT_ID },
            { "MicrosoftAppPassword", config.BOT_PASSWORD },
            { "MicrosoftAppTenantId", config.BOT_TENANT_ID },
        });
    })
    .ConfigureServices((hostContext, services) =>
    {
        var configuration = hostContext.Configuration;

        services.AddSingleton(hostContext.HostingEnvironment.ContentRootPath);

        // Create the Bot Framework Authentication to be used with the Bot Adapter.
        services.AddSingleton<BotFrameworkAuthentication, ConfigurationBotFrameworkAuthentication>();

        // Create the Cloud Adapter with error handling enabled.
        // Note: some classes expect a BotAdapter and some expect a BotFrameworkHttpAdapter, so
        // register the same adapter instance for all types.
        services.AddSingleton<CloudAdapter, AdapterWithErrorHandler>();
        services.AddSingleton<IBotFrameworkHttpAdapter>(sp => sp.GetService<CloudAdapter>());
        services.AddSingleton<BotAdapter>(sp => sp.GetService<CloudAdapter>());

        // Create the Conversation with notification feature enabled.
        services.AddSingleton(sp =>
        {
            var options = new ConversationOptions()
            {
                Adapter = sp.GetService<CloudAdapter>(),
                Notification = new NotificationOptions
                {
                    BotAppId = configuration["MicrosoftAppId"],
                },
            };

            return new ConversationBot(options);
        });

        // Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
        services.AddTransient<IBot, TeamsBot>();
    })
    .Build();

host.Run();