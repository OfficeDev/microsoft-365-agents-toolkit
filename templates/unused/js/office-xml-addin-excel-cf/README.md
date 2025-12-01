# Build Excel Custom Functions add-ins using Microsoft 365 Agents Toolkit

Excel add-ins are integrations built by third parties into Excel by using [Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) and [Office Platform capabilities](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins).

## Custom functions in Excel

Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.  

You can use this repository as a sample to base your own custom functions project from if you choose not to use the generator. For more detailed information about custom functions in Excel, see the [Custom functions overview](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-overview) article in the Office Add-ins documentation or see the [additional resources](#additional-resources) section of this repository.


## Prerequisites

- Node.js 16 or 18 (18 is preferred). Visit the  website to download and install the right version for your operating system.
- Office connected to a Microsoft 365 subscription (including Office on the web). If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program); for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/en-us/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/en-us/microsoft-365/buy/compare-all-microsoft-365-products).

## Run and Debug Excel Add-in

Before run and start the debug, make sure that:
1. Close all opened Office Application windows.
2. Click the *`Check and Install Dependencies`* in Microsoft 365 Agents Toolkit extension sidebar.

You can run and debug this project by either of the following ways:

- By hitting the `F5` key in Visual Studio Code.
- By clicking the *`Preview Your Add-in`* in Microsoft 365 Agents Toolkit extension sidebar.
- By running with command `npm run start` in the terminal.

## Debugging custom functions

This template supports debugging custom functions from [Visual Studio Code](https://code.visualstudio.com/). For more information see [Custom functions debugging](https://aka.ms/custom-functions-debug). For general information on debugging task panes and other Office Add-in parts, see [Test and debug Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).

## Understand the project

The add-in project that you've created contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the key files listed below. 

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/taskpane.css` file contains the CSS that's applied to content in the task pane.
- The `./src/taskpane/taskpane.js` file contains the Office JavaScript API code that facilitates interaction between the task pane and the Excel application.

## Validate manifest

You can check whether your manifest file is valid by either of the following ways:

- From Visual Studio Code: open Microsoft 365 Agents Toolkit extension sidebar and click *`Validate Manifest`*.
- From Terminal: run the command `npx --yes office-addin-manifest validate manifest.xml`

## Additional resources

- [Custom functions overview](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-overview)
- [Custom functions runtime](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime)
- [Custom functions troubleshoot](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-troubleshooting)
- [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- More Office Add-ins samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
