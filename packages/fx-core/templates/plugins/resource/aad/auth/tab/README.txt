Enable single sign-on for Teams tab applications
-------------------------

Files generated/updated in your project
-------------------------

1. New file - 'aad.template.json' is created in folder 'templates/appPackage'
   - The Microsoft Entra application manifest that is used to register the application with Microsoft Entra.
2. Update file - 'templates/appPackage/manifest.template.json'
   - An 'webApplicationInfo' object will be added into your app manifest template. This field is required by Teams when enabling SSO.
3. New file - 'Auth/tab'
   - Sample code and a 'README.txt' file. These files are provided for reference. See below for more information.
4. Update file - 'appsettings.json' and 'appsettings.Development.json'
   - Configs that will be used by TeamsFx SDK will be added into your app settings. Please update add the 'TeamsFx' object if you have other appsettings files.

Actions required - update your code to add SSO authentication
-------------------------

You need to update your application code to take advantage of the SSO authentication.

1. Move 'GetUserProfile.razor' file from 'Auth/tab' folder to 'Components/' folder.
   - 'GetUserProfile': This file implements a function that uses TeamsFx SDK to call Microsoft Graph API to get user info.
2. Rplace the 'AddSSO' component with 'GetUserProfile' component. To do this, just replace the following line: '<AddSSO />' with '<GetUserProfile />' in 'Components/Welcome.razor' file.

Debug your application
-------------------------

1. Right-click your project and select Microsoft 365 Agents Toolkit > Select Microsoft 365 Account
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
