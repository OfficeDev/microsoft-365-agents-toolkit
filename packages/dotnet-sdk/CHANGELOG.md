# 3.0.0-rc
- Breaking: Migrate dependencies from Bot Framework SDK to Microsoft Agents SDK.
- Breaking: Upgrade target framework to .NET 8.0.

# 2.5.0

- Update dependencies `Microsoft.Bot.Builder` and `Microsoft.Bot.Builder.Dialogs`.
- Update dependency `Microsoft.Identity.Client`.

# 2.4.1

- Update `@microsoft/teams-js` version to 2.22.0.

# 2.4.0

- Support to set notification local store file name using environment variable `TEAMSFX_NOTIFICATION_STORE_FILENAME`.

# 2.3.0

- Fix `NotificationMiddleware` reference key generation when bot is added to a Team.
- Update `botbuilder` dependencies

# 2.2.0

- Add `validationEnabled` to `GetPagedInstallationsAsync()` to enable or disable installation validation.
- Add new API `ValidateInstallationAsync` to validate bot installation.
- Deprecate the class `TeamsFx` and `MsGraphAuthProvider`. Please use `TokenCredential` and `GraphServiceClient(tokenCredential, new string[] { _scope })` instead to create Microsoft Graph Client instead.

# 2.1.0

- Add new interface `IConversationReferenceStore` to manage notification target references.
- Support to get a paginated list of targets where the bot is installed in notification bot.
- Support to get a paginated list of members of one-on-one, group, or team conversation in notification bot.
- Support to build `TeamsBotInstallation` instance with conversation reference.
- Deprecate the interface `INotificationTargetStorage` and the option `storage` in notification bot.
- Deprecate the API `GetMembersAsync`in notification bot.

# 2.0.0

- Support Graph SDK v5.6.0
- Remove `frameworkreference` and remove auth-start and auth-end html pages

# 1.2.0 -rc.2

- Update to use Teams JS SDK V2

# 1.2.0-rc

- Support adaptive card universal action handler in conversation bot.

# 1.1.0

- Support Bot SSO

# 1.0.0

- Version bump

# 0.5.0

- Add conversation SDK for command / notification bot

# 0.4.1-rc

- Integrate auth-start and auth-end html pages
- Remove "InitiateLoginEndpoint" option

# 0.4.0-rc

- Breaking: remove dependency of Microsoft.TeamsFx.SimpleAuth package

# 0.3.1

- Support .NET 6

# 0.2.0-rc

- Breaking: remove the dependency of TeamsFx JS SDK and implement in C#

# 0.1.0-rc

Initial release of the .NET SDK. Following features are included:

- TeamsUserCredential to simplify Team app authentication with Teams SSO support
- JS interop to call TeamsFx JS SDK
