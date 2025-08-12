// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.ComponentModel.DataAnnotations;

namespace {{YOUR_NAMESPACE}}.Configuration;

/// <summary>
/// Bot related authentication configuration.
/// </summary>
public class BotAuthenticationOptions
{
  /// <summary>
  /// The client (application) ID of an App Registration in the tenant.
  /// </summary>
  [Required(ErrorMessage = "Client id is required")]
  [RegularExpression(@"^[0-9A-Fa-f\-]{36}$")]
  public string ClientId { get; set; }

  /// <summary>
  /// The client (application) Secret of an App Registration in the tenant.
  /// </summary>
  [Required(ErrorMessage = "Client secret is required")]
  public string ClientSecret { get; set; }

  /// <summary>
  /// Authority URL that is used in OAuth On-behalf-of flow.
  /// </summary>
  [Required(ErrorMessage = "OAuth authority is required")]
  [RegularExpression(@"^http(s)?://[-a-zA-Z0-9@:%._\+~#=/]{1,100}$")]
  public string OAuthAuthority { get; set; }

  /// <summary>
  /// Application ID URI.
  /// </summary>
  [Required(ErrorMessage = "Application id uri is required")]
  public string ApplicationIdUri { get; set; }

  /// <summary>
  /// Login authentication start page endpoint.
  /// </summary>
  [Required(ErrorMessage = "Login authentication start page endpoint is required")]
  public string InitiateLoginEndpoint { get; set; }
}
