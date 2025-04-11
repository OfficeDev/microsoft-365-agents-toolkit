# Overview of the Dashboard template

This template showcases an app that embeds a canvas containing multiple cards that provide an overview of content in Microsoft Teams. Start with this template you can:

- Use widgets to display content from apps and services within your dashboard tab.
- Integrate your app with Graph API to visualize details about the implementation of the selected data.
- Create customizable dashboards that allow your business to set specific goals that help you track the information you need to view in multiple areas and across departments

## Get started with the Dashboard template

> **Prerequisites**
> To run the dashboard template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Set up your dev environment for extending apps across Microsoft 365](https://aka.ms/teamsfx-m365-apps-prerequisites)
>   Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.

**Congratulations**! You are running an application that can now show a dashboard in Teams:

![Dashboard](https://github.com/OfficeDev/TeamsFx/assets/107838226/9d0f4dcc-e216-418f-a947-671957f3dbee)

## What's included in the template

| Folder       | Contents                                            |
| ------------ | --------------------------------------------------- |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `src`        | The source code for the dashboard application |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                            |
| ------------------------------------ | --------------------------------------------------- |
| `src/services/chartService.js`       | A data retrieve implementation for the chart widget |
| `src/services/listService.js`        | A data retrieve implementation for the list widget  |
| `src/dashboards/SampleDashboard.jsx` | A sample dashboard layout implementation            |
| `src/styles/ChartWidget.css`         | The chart widget style file                         |
| `src/styles/ListWidget.css`          | The list widget style file                          |
| `src/widgets/ChartWidget.jsx`        | A widget implementation that can display a chart    |
| `src/widgets/ListWidget.jsx`         | A widget implementation that can display a list     |
| `src/App.css`                        | The style of application route                      |
| `src/App.jsx`                        | Application route                                   |

The following are project-related files. You generally will not need to customize these files.

| File                       | Contents                                   |
| -------------------------- | ------------------------------------------ |
| `src/main.css`             | The style of React application entry point |
| `src/main.jsx`             | The entry point for a React application    |
| `src/internal/context.jsx` | TeamsFx Context                            |
| `index.html`               | The entry point of the application         |
| `vite.config.js`           | Vite configuration file                    |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `m365agents.yml`       | This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `m365agents.local.yml` | This overrides `m365agents.yml` with actions that enable local execution and debugging.                                                     |

## Extend the Dashboard template to add a new widget

Follow the steps below to add a new widget to the dashboard:

1. [Step 1: Create a data retrive service](#step-1-create-a-data-retrive-service)
2. [Step 2: Create a widget file](#step-2-create-a-widget-file)
3. [Step 3: Add the widget to the dashboard](#step-3-add-the-widget-to-the-dashboard)

### Step 1: Create a data retrive service

Typically, a widget requires a service to retrieve the necessary data for displaying its content. This service can either fetch static data from a predefined source or retrieve dynamic data from a backend service or API.

For instance, we will implement a service that returns static data and place it under the `src/services` directory.

Here is a sample service for retrieving static data:

```javascript
//sampleService.js
export const getSampleData = () => {
  return { content: "Hello world!" };
};
```

### Step 2: Create a widget file

Create a widget file in the `src/widgets` folder. Inherit the `BaseWidget` class from `@microsoft/teamsfx-react`. The following table lists the methods that you can override to customize your widget.

| Methods     | Function                                                                                                                                      |
| ----------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `getData()` | This method is used to get the data for the widget. You can implement it to get data from the backend service or from the Microsoft Graph API |
| `header()`  | Customize the content of the widget header                                                                                                    |
| `body()`    | Customize the content of the widget body                                                                                                      |
| `footer()`  | Customize the content of the widget footer                                                                                                    |
| `styling()` | Customize the widget style                                                                                                                    |

> All method overrides are optional.

Here's a sample widget implementation:

```javascript
//SampleWidget.jsx
import { Button, Text } from "@fluentui/react-components";
import { BaseWidget } from "@microsoft/teamsfx-react";
import { getSampleData } from "../services/sampleService";

export class SampleWidget extends BaseWidget {
  async getData() {
    return getSampleData();
  }

  header() {
    return <Text>Sample Widget</Text>;
  }

  body() {
    return <div>{this.state.data?.content}</div>;
  }

  footer() {
    return <Button>View Details</Button>;
  }
}
```

### Step 3: Add the widget to the dashboard

Open the `src/dashboards/SampleDashboard.jsx` file and add the widget to the implementation of the `layout` method. If you want to create a new dashboard, please refer to [How to add a new dashboard](https://aka.ms/teamsfx-dashboard-new#how-to-add-a-new-dashboard).

```jsx
layout() {
  return (
    <>
      <ListWidget />
      <ChartWidget />
      <SampleWidget />
    </>
  );
}
```

Optional: If you want to arrange multiple widgets in the same column, you can refer to the following code snippet:

```css
.one-column {
  display: grid;
  gap: 20px;
  grid-template-rows: 1fr 1fr;
}
```

```jsx
layout() {
  return (
    <>
      <ListWidget />
      <div className="one-column">
        <ChartWidget />
        <SampleWidget />
      </div>
    </>
  );
}
```

Congratulations, you've just added your own widget! To learn more about the dashboard template, [visit the documentation](https://aka.ms/teamsfx-dashboard-new). You can find more scenarios like:

- [Customize the widget](https://aka.ms/teamsfx-dashboard-new#customize-the-widget)
- [Customize the dashboard layout](https://aka.ms/teamsfx-dashboard-new#customize-the-dashboard-layout)
- [Create a data loader](https://aka.ms/teamsfx-dashboard-new#how-to-include-a-data-loader)
- [Refresh data based on the schedule](https://aka.ms/teamsfx-dashboard-new#how-to-refresh-data-as-scheduled)
- [Handle empty state](https://aka.ms/teamsfx-dashboard-new#how-to-handle-empty-state)
- [Add a new dashboard](https://aka.ms/teamsfx-dashboard-new#how-to-add-a-new-dashboard)
- [Use Microsoft Graph Toolkit as widget content](https://aka.ms/teamsfx-dashboard-new#how-to-use-microsoft-graph-toolkit-as-widget-content)
- [Embed Power BI to dashboard](https://aka.ms/teamsfx-dashboard-new#how-to-embed-power-bi-to-dashboard)
- [How to add a new Graph API call](https://aka.ms/teamsfx-dashboard-new#how-to-add-a-new-graph-api-call)
- [Enable the app for multi-tenant](https://github.com/OfficeDev/TeamsFx/wiki/Multi-tenancy-Support-for-Azure-AD-app)
- [Preview the app on mobile clients](https://aka.ms/teamsfx-mobile)

## Additional resources

- [Fluent UI](https://react.fluentui.dev/?path=/docs/concepts-introduction--page)
- [Fluent UI React Charting Example](https://fluentuipr.z22.web.core.windows.net/heads/master/react-charting/demo/index.html#/)
