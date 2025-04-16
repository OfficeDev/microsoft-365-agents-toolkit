# Enable single sign-on for tab applications

Microsoft Teams lets your app obtain the signed-in user token to access Microsoft Graph and other APIs. The Microsoft 365 Agents Toolkit simplifies this by wrapping Microsoft Entra ID flows in easy-to-use APIs, making it simple to add SSO features to your Microsoft Teams App.

# Changes to your project

When you added the SSO feature to your application, Microsoft 365 Agents Toolkit updated your project to support SSO:

After you successfully added SSO into your project, Microsoft 365 Agents Toolkit will create and modify some files that helps you implement SSO feature.

| Action | File | Description |
| - | - | - |
| Create| `aad.template.json` under `templates/appPackage` | The Microsoft Entra application manifest that is used to register the application with Microsoft Entra. |
| Modify | `manifest.template.json` under `templates/appPackage` | An `webApplicationInfo` object will be added into your app manifest template. This field is required by Teams when enabling SSO. |
| Create | `auth/tab` | Reference code, redirect pages and a `README.md` file. These files are provided for reference. See below for more information. |

# Update your code to add SSO

The Microsoft 365 Agents Toolkit has configured your app for SSO, but you'll need to update your business logic to fully utilize this feature.

1. Move `auth-start.html` and `auth-end.html` in `auth/tab/public` folder to `tabs/public/`.
These two HTML files are used for auth redirects.

1. Move `sso` folder under `auth/tab` to `tabs/src/sso/`.

    `InitTeamsFx`: This file implements a function that initialize TeamsFx SDK and will open `GetUserProfile` component after SDK is initialized.

    `GetUserProfile`: This file implements a function that calls Microsoft Graph API to get user info.

2. Add the following lines to `tabs/src/components/sample/Welcome.*` to import `InitTeamsFx`:
    ```
    import { InitTeamsFx } from "../../sso/InitTeamsFx";
    ```
3. Replace the following line: `<AddSSO />` with `<InitTeamsFx />` to replace the `AddSSO` component with `InitTeamsFx` component.

# Debug your application

You can debug your application by pressing F5.

Microsoft 365 Agents Toolkit will use the Microsoft Entra manifest file to register a Microsoft Entra application registered for SSO.

To learn more about Microsoft 365 Agents Toolkit local debug functionalities, refer to this [document](https://docs.microsoft.com/microsoftteams/platform/toolkit/debug-local).

# Customize Microsoft Entra applications

The Microsoft Entra [manifest](https://docs.microsoft.com/azure/active-directory/develop/reference-app-manifest) allows you to customize various aspects of your application registration. You can update the manifest as needed.

Follow this [document](https://aka.ms/teamsfx-aad-manifest#how-to-customize-the-aad-manifest-template) if you need to include additional API permissions to access your desired APIs.

Follow this [document](https://aka.ms/teamsfx-aad-manifest#How-to-view-the-AAD-app-on-the-Azure-portal) to view your Microsoft Entra application in Azure Portal.
