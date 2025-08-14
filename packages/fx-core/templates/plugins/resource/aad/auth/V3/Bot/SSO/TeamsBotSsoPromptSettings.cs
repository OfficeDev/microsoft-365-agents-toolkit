// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using {{YOUR_NAMESPACE}}.Configuration;

namespace {{YOUR_NAMESPACE}}.SSO;

/// <summary>
/// Contains settings for an <see cref="TeamsBotSsoPrompt"/>.
/// </summary>
public class TeamsBotSsoPromptSettings
{

  /// <summary>
  /// Constructor of TeamsBotSsoPromptSettings
  /// </summary>
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
  public string[] Scopes { get; set; }

  /// <summary>
  /// Gets or sets the number of milliseconds the prompt waits for the user to authenticate.
  /// Default is 900,000 (15 minutes).
  /// </summary>
  /// <value>The number of milliseconds the prompt waits for the user to authenticate.</value>
  public int Timeout { get; set; }

  /// <summary>
  /// Gets or sets bot related authentication options.
  /// </summary>
  public BotAuthenticationOptions BotAuthOptions { get; set; }
}