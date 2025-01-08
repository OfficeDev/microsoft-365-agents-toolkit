using {{SafeProjectName}};
using Microsoft.Agents.Authentication;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Protocols.Connector;
using Microsoft.Agents.Protocols.Primitives;

namespace Microsoft.Agents.Hosting.Setup
{
    public static class ServiceCollectionExtensions
    {
        public static IHostApplicationBuilder AddBot<T, TImpl>(this IHostApplicationBuilder builder)
            where T : IBot
            where TImpl : class, T
        {
            // Add Connections object to access configured token connections.
            builder.Services.AddSingleton<IConnections, ConfigurationConnections>();

            // Add factory for ConnectorClient and UserTokenClient creation
            builder.Services.AddSingleton<IChannelServiceClientFactory, RestChannelServiceClientFactory>();

            // Add the CloudAdapter with customized error handling enabled.
            builder.Services.AddCloudAdapter<AdapterWithErrorHandler>();

            // Add the Bot,  this is the primary worker for the bot. 
            builder.Services.AddTransient<IBot, TImpl>();

            return builder;
        }
    }
}
