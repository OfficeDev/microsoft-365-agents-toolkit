# Review project upgrade and changes

Your project has been automatically upgraded to the latest version of the Microsoft 365 Agents Toolkit. This upgrade is irreversible; to revert to the previous version, follow the [rollback instructions](#how-to-roll-back-after-the-upgrade).

> If you encountered a problem after the automatic upgrade, review the [known issues](https://aka.ms/teams-toolkit-5.0-upgrade#known-issues). If you are still having trouble, contact ttkfeedback@microsoft.com or [file an issue on GitHub](https://github.com/officedev/teamsfx/issues) with details about the problem.

## Important changes to your development flow

The new features in this release make developing your apps simpler and more flexible, but also include changes to how the toolkit creates projects, automates configuration, and deploys resources. Visit [the upgrade guide](https://aka.ms/teamsfx-v5.0-guide) for more info about all of the changes.

Some of the changes you may immediately notice are:

* Configuration of the lifecycle management using Provision, Deploy, and Publish are now fully customizable and expressed in `m365agents.local.yml` and `m365agents.yml`. [More info](https://aka.ms/teamsfx-v5.0-guide#project-files)
* Configuration and values that were saved in `.fx/config` and `.fx/state` are now handled with environment files and saved to `/env` by default. [More info](https://aka.ms/teams-toolkit-5.0-upgrade#environment-management)
* The changes to use environment files give greater flexibility of configuring which resources are used to provision, but may require some manual steps when creating new environments. [More info](https://aka.ms/teams-toolkit-5.0-upgrade#environment-management)

If you are using Visual Studio version 17.7 or later and developing a bot app locally:
* You can use the dev tunnel as your tunneling service. [More Info](https://aka.ms/vs-dev-tunnel-guidance)
* In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel.

If you're using APIM:
* You need to provide values to `APIM__PUBLISHEREMAIL` and the `APIM__PUBLISHERNAME` environment variables. [More info](https://aka.ms/teams-toolkit-5.0-upgrade#provision-apim-service)

You can [view these changes on GitHub](https://aka.ms/teams-toolkit-5.0-upgrade#feature-changes-that-impact-your-development-flow).

## Changes to your project file structure

1. Created `m365agents.yml` and `m365agents.local.yml` in your project's root directory.
2. Moved environment files in `.fx` to `.env.{env}` in `env` folder.
3. If your project contains file `.fx/states/{env}.userdata`, the content will be moved to `.env.{env}.user` in `env` folder
4. Moved `templates/appPackage` to `appPackage`, renamed `manifest.template.json` to `manifest.json` and placeholders in it will be updated using the latest default conventions.
5. If your project contains `templates/appPackage/aad.template.json`, it will be moved and renamed to `aad.manifest.json` and the templated variables names are updated to the latest default conventions.
6. If your project contains file `.vscode/tasks.json` and `.vscode/launch.json`, they will be updated.
7. Updated `.gitignore` to ignore new environment user files.
8. Removed `.fx` folder.

You can [view these changes on GitHub](https://aka.ms/teams-toolkit-5.0-upgrade#file-changes).

## Known issues

1. If your project only contains a bot, you might get an error about `STATE__FX_RESOURCE_FRONTEND_HOSTING__ENDPOINT` missing when running Provision or using Start Debugging. Find this placeholder variable in `appPackage/manifest.json` and replace it with a valid URL to resolve this issue. [More Info](https://aka.ms/teams-toolkit-5.0-upgrade#state__fx_resource_frontend_hosting__endpoint-missing-error-in-some-projects)
2. If your project is created with Visual Studio version < 17.4, you might get an error like `InvalidParameter: Following parameter is missing or invalid for aadApp/create action: name` when running commands. [Try these steps](#how-to-roll-back-after-the-upgrade) to roll back the changes, install VS 17.4, and run the upgrade again.
3. If your tab app is created with Microsoft 365 Agents Toolkit 3.2.0 or an earlier version, you may get an error like `simpleAuthEndpoint in configuration is invalid` when remote debugging your app. [Try these steps](https://aka.ms/teams-toolkit-5.0-upgrade#simpleauthendpoint-in-configuration-is-invalid) to learn how to mitigate this error.
4. If your project was successfully provisioned before, but after upgrading it cannot be provisioned or published using the `teamsApp/validateAppPackage` action, try using the [validation report in Teams Developer Portal](https://dev.teams.microsoft.com/validation) to check the manifest for errors. [More Info](https://aka.ms/teams-toolkit-5.0-upgrade#teamsappvalidateapppackage-failed-error)

## How to roll back after the upgrade

Follow these steps if you want to restore your project configuration after the upgrade is successful or need to use a previous version of Microsoft 365 Agents Toolkit:
1. Copy everything in the `.backup` folder that was generated during the upgrade to your project root folder.
2. Delete the new files created during the upgrade. The [Changes to your project](#changes-to-your-project) section contains info on everything that was created.

You can [view these steps on GitHub](https://aka.ms/teams-toolkit-5.0-upgrade#how-to-roll-back).
