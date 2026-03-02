// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Identity.Client;

namespace {{YOUR_NAMESPACE}}.SSO
{
  /// <summary>
  /// Helper class used to simplify unit test.
  /// https://stackoverflow.com/questions/65334284/how-to-mock-moq-iconfidentialclientapplication-which-has-sealed-setup-abstract
  /// </summary>
  internal class IdentityClientAdapter : IIdentityClientAdapter
  {
    private readonly IConfidentialClientApplication _confidentialClientApplication;

    public IdentityClientAdapter(IConfidentialClientApplication confidentialClientApplication)
    {
      _confidentialClientApplication = confidentialClientApplication;
    }

    public async Task<AuthenticationResult> GetAccessToken(string ssoToken, IEnumerable<string> scopes)
    {
      var userAssertion = new UserAssertion(ssoToken);
      return await _confidentialClientApplication
          .AcquireTokenOnBehalfOf(scopes, userAssertion)
          .ExecuteAsync()
          .ConfigureAwait(false);
    }
  }
}
