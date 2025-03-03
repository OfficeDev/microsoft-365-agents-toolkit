# Welcome to Teams Toolkit!

## Quick Start

> **Prerequisites**
>
> To run this app template in your local dev machine, you will need:
>
> - [Visual Studio 2022](https://aka.ms/vs) 17.9 or higher and [install Teams Toolkit](https://aka.ms/install-teams-toolkit-vs)
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts).

1. Right-click your project and select `Teams Toolkit > Provision in the Cloud..`. You can find everything it will do in the `teamsapp.yml`.
2. If prompted, sign in with a Microsoft 365 account for the Teams organization you want 
to install the app to.
3. Right-click your project and select `Teams Toolkit > Preview in > Teams`.
4. To trigger the Message Extension, you can click the `+` under compose message area to find your message extension.
   > Note: Please make sure to switch to New Teams when Teams web client has launched

{{#ApiKey}}
> [!NOTE]
> Teams Toolkit will ask you for your API key during provision. The API key will be securely stored with [Teams Developer Portal](https://dev.teams.microsoft.com/home) and used by Teams client to access your API in runtime. Teams Toolkit will not store your API key.
{{/ApiKey}}

{{#OAuth}}
> [!NOTE]
> If your identity server needs Proof of Key Code Exchange (PKCE) for token exchange, uncomment the `isPKCEEnabled` property in the` oauth/register` section of the `teamsapp.yml` file shown as below:
```yaml
  - uses: oauth/register
    with:
      name: {{ApiSpecAuthName}}
      flow: authorizationCode
      # Teams app ID
      appId: ${{TEAMS_APP_ID}}
      # Path to OpenAPI description document
      apiSpecPath: {{{ApiSpecPath}}}
      # Uncomment below property to use proof key for code exchange (PKCE)
      isPKCEEnabled: true
    writeToEnvironmentFile:
      configurationId: {{ApiSpecAuthRegistrationIdEnvName}}
```
> Teams Toolkit will ask you for your Client ID and Client Secret for Oauth2 during provision. These information will be securely stored with [Teams Developer Portal](https://dev.teams.microsoft.com/home) and used by Teams client to access your API in runtime. Teams Toolkit will not store your Client ID and Client Secret.
{{/OAuth}}

## Learn more

- [Extend Teams platform with APIs](https://aka.ms/teamsfx-api-plugin)

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem.
Or, you can create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues
