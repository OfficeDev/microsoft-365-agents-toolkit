# Build PowerPoint add-ins using Microsoft 365 Agents Toolkit

PowerPoint add-ins are integrations built by third parties into PowerPoint by using [PowerPoint JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/powerpoint-add-ins-reference-overview) and [Office Platform capabilities](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins).

## Prerequisites

- Node.js 16 or 18 (18 is preferred). Visit the  website to download and install the right version for your operating system.
- Office connected to a Microsoft 365 subscription (including Office on the web). If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program); for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/en-us/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/en-us/microsoft-365/buy/compare-all-microsoft-365-products).

## Instructions

Before run and start the debug, make sure that:
1. Close all opened Office Application windows.
2. Click the *`Check and Install Dependencies`* in Microsoft 365 Agents Toolkit extension sidebar.

- Run the following command to configure single-sign on for your add-in project.

```shell
npm run configure-sso
```

- A web browser window will open and prompt you to sign in to Azure. Sign in to Azure using your Microsoft 365 administrator credentials. These credentials will be used to register a new application in Azure and configure the settings required by SSO.

- Build the project, start the local web server, and side-load your add-in in the previously selected Office client application by either of the following ways:
  - By hitting the `F5` key in Visual Studio Code.
  - By clicking the *`Preview Your Add-in`* in Microsoft 365 Agents Toolkit extension sidebar.
  - By running with command `npm run start` in the terminal.

> [!NOTE]
> Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provide.

- In the Office client application that opens when you run the previous command, make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the previous section. Doing so establishes the appropriate conditions for SSO to succeed.

- In the Office client application, choose the Home tab, and then choose the Show Task Pane button in the ribbon to open the add-in task pane.

- At the bottom of the task pane, choose the Get My User Profile Information button to initiate the SSO process.

> If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication. This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account. Choose the Accept button in the dialog window to continue.

- The add-in retrieves profile information for the signed-in user and writes it to the document.

## Understand the project

The add-in project that you've created contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the key files listed below. 

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/taskpane.css` file contains the CSS that's applied to content in the task pane.
- The `./src/taskpane/taskpane.ts` file contains the Office JavaScript API code that facilitates interaction between the task pane and the PowerPoint application.

## Validate manifest

You can check whether your manifest file is valid by either of the following ways:

- From Visual Studio Code: open Microsoft 365 Agents Toolkit extension sidebar and click *`Validate Manifest`*.
- From Terminal: run the command `npx --yes office-addin-manifest validate manifest.xml`