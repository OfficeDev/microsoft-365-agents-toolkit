# 3.0.0
- Remove deprecated classes and APIs: `TeamsFx`, `CommandBot`, `CardActionBot`, `CommandBot`, Microsoft Graph Client, etc.
- Update dependencies to support Node.js v20.

# 2.3.2
- Update peer dependency `@microsoft/teams-js` version to `^2.19.0`.

# 2.3.0

- Add an optional parameter `validationEnabled` to `getPagedInstallations` to enable or disable installation validation.
- The `TeamsFx`, `AuthenticationConfiguration`, and `handleMessageExtensionQueryWithSSO` are deprecated and will be removed in a future release.
- The `MSGraphAuthProvider` class, `createMicrosoftGraphClient` and `createMicrosoftGraphClientWithCredential` are deprecated. Use the `TokenCredentialAuthenticationProvider` class and the `Client.InitialWithMiddleware` method from the `@microsoft/microsoft-graph-client` package instead.

# 2.2.0

- Deprecate interfaces and APIs related with `BotFrameworkAdapter`.
- Add module `BotBuilderCloudAdapter` and migrate interfaces and APIs to use `CloudAdapter`.

# 2.0.0

- Refactor SDK and make it easy to use (Docs will be ready in next release).
- Support Teams JS SDK 2.0.
- Add error details when projects not running inside Teams.
- `getTediousConnectionConfig` is now deprecated. We recommend you compose your own Tedious configuration for better flexibility.

# 0.6.0

- Breaking: Remove loadConfiguration() function and add TeamsFx class
- Support automatic choosing of credential class based on IdentityType

# 0.5.0

- Breaking: TeamsUserCredential now will use Auth Code Flow with PKCE for SPA authentication. You can find more detail here: https://aka.ms/teamsfx-auth-code-flow.
- Support multiple SQL database connection
- Fix TeamsBotSsoPrompt runtime error

# 0.4.1

- Move "@microsoft/teams-js" to peer dependency

# 0.3.0

- Support certificate-based authentication for TeamsFx SDK

# 0.0.5

- Read SQL database name from `SQL_DATABASE_NAME` environment variable
- Add more error types for credential classes

# 0.0.4

Initial alpha release of the SDK. Following features are included:

- Credentials to simplify Team app authentication with Teams SSO support
- Simplify authentication to Microsoft Graph APIs
- TeamsBotSsoPrompt to support Teams Bot development with SSO
- DefaultTediousConnectionConfiguration to connect to MSSQL database
