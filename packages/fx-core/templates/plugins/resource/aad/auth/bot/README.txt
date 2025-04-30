Enable single sign-on for Teams bot applications
-------------------------

For Teams bot application, SSO manifests as an Adaptive Card which the user can interact with to invoke the Microsoft Entra consent flow.

Files generated/updated in your project
-------------------------

1. New file - `aad.template.json` is created in folder `templates/appPackage`
   - The Azure Active Directory application manifest that is used to register the application with Microsoft Entra.
2.  Update file - 'templates/appPackage/manifest.template.json'
   - An `webApplicationInfo` object will be added into your app manifest template. This field is required by Teams when enabling SSO. |
3. New file - `Auth/bot`
    - Sample code, redirect pages and a `README.txt` file. These files are provided for reference. See below for more information. |
4. Update file - 'appsettings.json' and 'appsettings.Development.json'
   - Configs that will be used by TeamsFx SDK will be added into your app settings. Please update add the 'TeamsFx' object if you have other appsettings files.

Actions required - update your code to add SSO authentication
-------------------------
Note: This part is for `command and response bot`.

1. Please upgrade your SDK and make sure your SDK version:
   TeamsFx: >= 1.1.0
   Microsoft.Bot.Builder >= 4.17.1

2. Create "Pages" folder and move files in `Auth/bot/Pages` folder to `Pages`
   `Auth/bot/Pages` folder contains HTML pages that hosted by bot application. When single sign-on flows are initiated with Microsoft Entra ID, Microsoft Entra ID will redirect the user to these pages.

3. Create "SSO" folder and move files in 'Auth/bot/SSO' folder to 'SSO'
   This folder contains two files as reference for SSO implementation:
   2.1 SsoDialog.cs: This creates a ComponentDialog that used for SSO.
   2.2 TeamsSsoBot.cs: This create a TeamsActivityHandler with `SsoDialog` and add 'showUserInfo' as a command that can be triggered.
   2.3 SsoOperations.cs: This implements class with a function to get user info with SSO token. You can follow this method and create your own method that requires SSO token.
   Note: Remember to replace '{Your_NameSpace}' with your project namespace.

4. Update 'Program.cs'
    4.1 Find code: 'builder.Services.AddSingleton<BotFrameworkAuthentication, ConfigurationBotFrameworkAuthentication>();'
        and add the following code below:
        '''
        builder.Services.AddRazorPages();

        // Create the Bot Framework Adapter with error handling enabled.                                        
        builder.Services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();

        builder.Services.AddSingleton<IStorage, MemoryStorage>();
        // Create the Conversation state. (Used by the Dialog system itself.)
        builder.Services.AddSingleton<ConversationState>();

        // The Dialog that will be run by the bot.
        builder.Services.AddSingleton<SsoDialog>();

        // Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
        builder.Services.AddTransient<IBot, TeamsSsoBot<SsoDialog>>();

        builder.Services.AddOptions<AuthenticationOptions>().Bind(builder.Configuration.GetSection("TeamsFx").GetSection(AuthenticationOptions.Authentication)).ValidateDataAnnotations();
        builder.Services.AddOptions<BotAuthenticationOptions>().Configure<IOptions<AuthenticationOptions>>((botAuthOption, authOptions) => {
            AuthenticationOptions authOptionsValue = authOptions.Value;
            botAuthOption.ClientId = authOptionsValue.ClientId;
            botAuthOption.ClientSecret = authOptionsValue.ClientSecret;
            botAuthOption.OAuthAuthority = authOptionsValue.OAuthAuthority;
            botAuthOption.ApplicationIdUri = authOptionsValue.ApplicationIdUri;
            botAuthOption.InitiateLoginEndpoint = authOptionsValue.Bot.InitiateLoginEndpoint;
        }).ValidateDataAnnotations();
        '''
    4.2 Find the following lines:
        '''
        builder.Services.AddSingleton<HelloWorldCommandHandler>();
        builder.Services.AddSingleton(sp =>
        {
          var options = new ConversationOptions()
          {
            Adapter = sp.GetService<CloudAdapter>(),
            Command = new CommandOptions()
            {
              Commands = new List<ITeamsCommandHandler> { sp.GetService<HelloWorldCommandHandler>() }
            }
          };

          return new ConversationBot(options);
        });
        '''
        and replace with:
        '''
        builder.Services.AddSingleton(sp =>
        {
          var options = new ConversationOptions()
          {
            Adapter = sp.GetService<CloudAdapter>(),
            Command = new CommandOptions()
            {
              Commands = new List<ITeamsCommandHandler> { }
            }
          };

          return new ConversationBot(options);
        });
        '''
    4.3 Find and delete the following code:
        '''
        // Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
        builder.Services.AddTransient<IBot, TeamsBot>();
        '''
    4.4 Find the following code:
        '''
        app.UseEndpoints(endpoints =>
        {
          endpoints.MapControllers();
        });
        '''
        and replace with:
        '''
        app.UseEndpoints(endpoints =>
        {
          endpoints.MapControllers();
          endpoints.MapRazorPages();
        });
        '''

5. Register your command in the app manifest. Open 'Templates/appPackage/manifest.template.json', and add following lines under `commands` in `commandLists` of your bot:
    '''
    {
      "title": "show",
      "description": "Show user profile using Single Sign On feature"
    }
    '''

(Optional) Add a new command to the bot
-------------------------
After successfully add SSO in your project, you can also add a new command.
1. Create a new method in class SsoOperations in 'SSO/SsoOperations' and add your own business logic to call Graph API:
  '''
  public static async Task GetUserImageInfo(ITurnContext stepContext, string token, BotAuthenticationOptions botAuthOptions)
  {
      await stepContext.SendActivityAsync("Retrieving user information from Microsoft Graph ...");
      var authProvider = new DelegateAuthenticationProvider((request) =>
      {
          request.Headers.Authorization =
              new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
          return Task.CompletedTask;
      });
      var graphClient = new GraphServiceClient(authProvider);

      // You can add following code to get your photo size:
      // var photo = await graphClient.Me.Photo.Request().GetAsync();
      // await stepContext.SendActivityAsync($"Size of your photo is: {photo.Width} * {photo.Height}");
  }
  '''

2. Register a new command using 'addCommand' in 'TeamsSsoBot':
  Find the following line:
  '''
  ((SsoDialog)_dialog).addCommand("showUserInfo", "show", SsoOperations.ShowUserInfo);
  '''
  and add following lines after the above line to register a new command 'photo' and hook up with method 'GetUserImageInfo' added above:
  '''
  ((SsoDialog)_dialog).addCommand("getUserImageInfo", "photo", SsoOperations.GetUserImageInfo);
  '''

3. Register your command in the app manifest. Open 'Templates/appPackage/manifest.template.json', and add following lines under `commands` in `commandLists` of your bot:
    '''
    {
      "title": "photo",
      "description": "Show user photo size using Single Sign On feature"
    }
    '''

Debug your application
-------------------------
You can debug your application by:

1. Right-click your project and select Microsoft 365 Agents Toolkit > Prepare Teams app dependencies
2. If prompted, sign in with an M365 account for the Teams organization you want 
to install the app to
3. Press F5, or select the Debug > Start Debugging menu in Visual Studio
4. In the launched browser, select the Add button to load the app in Teams

Microsoft 365 Agents Toolkit will use the Microsoft Entra manifest file to register a Microsoft Entra application registered for SSO.

To learn more about Microsoft 365 Agents Toolkit local debug functionalities, refer to https://docs.microsoft.com/microsoftteams/platform/toolkit/debug-local.

Customize Microsoft Entra applications
-------------------------
The Microsoft Entra manifest allows you to customize various aspects of your application registration. You can update the manifest as needed.
Related Doc: https://docs.microsoft.com/azure/active-directory/develop/reference-app-manifest

Follow https://aka.ms/teamsfx-aad-manifest#how-to-customize-the-aad-manifest-template if you need to include additional API permissions to access your desired APIs.

Follow https://aka.ms/teamsfx-aad-manifest#How-to-view-the-AAD-app-on-the-Azure-portal to view your Microsoft Entra application in Azure Portal.
