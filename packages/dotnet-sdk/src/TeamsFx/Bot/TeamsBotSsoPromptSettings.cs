// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.TeamsFx.Configuration;

namespace Microsoft.TeamsFx.Bot;

/// <summary>
/// Contains settings for an <see cref="TeamsBotSsoPrompt"/>.
/// </summary>
[Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
public class TeamsBotSsoPromptSettings
{

    /// <summary>
    /// Constructor of TeamsBotSsoPromptSettings
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public TeamsBotSsoPromptSettings(BotAuthenticationOptions botAuthOptions, string[] scopes, int timeout = 900000)
    {
        BotAuthOptions = botAuthOptions;
        Scopes = scopes;
        Timeout = timeout;
    }

    /// <summary>
    /// Gets or sets the array of strings that declare the desired permissions and the resources requested.
    /// </summary>
    /// <value>The array of strings that declare the desired permissions and the resources requested.</value>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public string[] Scopes { get; set; }

    /// <summary>
    /// Gets or sets the number of milliseconds the prompt waits for the user to authenticate.
    /// Default is 900,000 (15 minutes).
    /// </summary>
    /// <value>The number of milliseconds the prompt waits for the user to authenticate.</value>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public int Timeout { get; set; }

    /// <summary>
    /// Gets or sets bot related authentication options.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public BotAuthenticationOptions BotAuthOptions { get; set; }
}