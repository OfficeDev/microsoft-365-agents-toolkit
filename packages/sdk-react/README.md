# TeamsFx SDK for React

> This package will be deprecated by 2026-07. You can copy hooks to your Teams app source code.

TeamsFx SDK provides [React hooks](https://reactjs.org/docs/hooks-intro.html) to reduce the developer tasks of integrating TeamsFx with React and leverage Teams SSO.

Use the library to:

- Call Graph API using an authenticated client.
- Customize TeamsFx easily in React app.

[Source code](https://github.com/OfficeDev/TeamsFx/tree/main/packages/sdk-react) |
[Package (NPM)](https://www.npmjs.com/package/@microsoft/teamsfx-react) |
[Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Getting started

> Important: Please be advised that access tokens are stored in sessionStorage for you by default. This can make it possible for malicious code in your app (or code pasted into a console on your page) to access APIs at the same privilege level as your client application. Please ensure you only request the minimum necessary scopes from your client application and perform any sensitive operations from server-side code that your client has to authenticate with.

TeamsFx SDK and React hooks are pre-configured in scaffolded project using Teams Toolkit extension for Visual Studio and vscode, or the `teamsfx` cli from the `teamsfx-cli` npm package.
Please check the [README](https://github.com/OfficeDev/TeamsFx/blob/main/packages/vscode-extension/README.md) to see how to create a Teams App project.

### Prerequisites

- Node.js version 18.x.x or higher
- PNPM version 8.x.x or higher
- TeamsFx SDK version 0.6.0 or higher 
- A project created by the Teams Toolkit VS Code extension or `teamsfx` CLI tool.

### Install the `@microsoft/teamsfx-react` package

Install the TeamsFx SDK for TypeScript/JavaScript with `npm`:

```bash
npm install @microsoft/teamsfx-react
```

Please also install the peer dependencies if you are using npm 6. 
```bash
npm install @microsoft/teamsfx@^2.0.0 @microsoft/teams-js@^2.0.0 react@^18.2.0 react-dom@^18.2.0 @fluentui/react-components@^9.16.0 @microsoft/microsoft-graph-client@^3.0.1
```

#### Note

For `@microsoft/teamsfx-react@^3.0.0`, it migrated to use `@fluentui/react-components` and did not support `@fluentui/react-northstar` any longer. If you still wish to use `@fluentui/react-northstar`, please use `@microsoft/teamsfx-react@<3.0.0`. To install the peer dependencies, you could use the following line:

```bash
npm install @microsoft/teamsfx@^2.0.0 @microsoft/teams-js@^2.0.0 react@^16.8.6 react-dom@^16.8.6 @fluentui/react-northstar@^0.62.0 @microsoft/microsoft-graph-client@^3.0.1
```

For `@microsoft/teamsfx-react@^2.0.0` and `@microsoft/teamsfx-react@^3.0.0`, they depend on `@microsoft/teamsfx@^2.0.0` and `@microsoft/teams-js@^2.0.0`. If you wish to use lower versions, please install the peer dependencies as follows:

```bash
npm install @microsoft/teamsfx@^0.6.0 react@^16.8.6 @fluentui/react-northstar@^0.60.1 msteams-react-base-component@^3.1.1
```

### Scenario

TeamsFx SDK for React is built to be used in React application. You can develop a new react web app for Teams Tab scenario.

### Calling the Microsoft Graph API

The SDK provides custom React hook `useGraph()` that provides an authenticated Graph client instance. Please use this hook to call Microsoft Graph API.

#### 1. Implement business logic and specify resource scope

Use the snippet below:

```ts
const { loading, error, data, reload } = useGraph(
    async (graph, teamsfx, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();

      let photoUrl = "";
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, photoUrl };
    },
    { scope: ["User.Read"] }
  );
```

> Note: `useGraph()` has been deprecated and will be removed in the future release. Please use `useGraphWithCredential()` instead as below:

```ts
const { loading, error, data, reload } = useGraphWithCredential(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();

      let photoUrl = "";
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, photoUrl };
    },
    { scope: ["User.Read"] }
  );
```

#### 2. Render Graph data with React

You can bind the `reload` function with a button to refresh data on demand.

```ts
return (
  <div>
    <h3>Example: Get the user's profile</h3>
    <div className="section-margin">
      <p>Click below to authorize button to grant permission to using Microsoft Graph.</p>
      <Button primary content="Authorize" disabled={loading} onClick={reload} />
      <PersonCardFluentUI loading={loading} data={data} error={error} />
    </div>
  </div>
);
```

### Building a Dashboard Tab

#### 1. Create a new widget

Here is an example of creating a new widget:

```tsx
import { Button, Text } from "@fluentui/react-components";
import { BaseWidget } from "@microsoft/teamsfx-react";
import { SampleModel } from "../models/sampleModel";
import { getSampleData } from "../services/sampleService";

interface SampleWidgetState {
  data?: SampleModel;
}

export class SampleWidget extends BaseWidget<any, SampleWidgetState> {
  override async getData(): Promise<SampleWidgetState> {
    return { data: getSampleData() };
  }

  override header(): JSX.Element | undefined {
    return <Text>Sample Widget</Text>;
  }

  override body(): JSX.Element | undefined {
    return <div>{this.state.data?.content}</div>;
  }

  override footer(): JSX.Element | undefined {
    return <Button>View Details</Button>;
  }
}
```

#### 2. Create a new dashboard

Here is an example of creating a new dashboard:

```tsx
import { BaseDashboard } from "@microsoft/teamsfx-react";
import ListWidget from "../widgets/ListWidget";
import ChartWidget from "../widgets/ChartWidget";

export default class YourDashboard extends BaseDashboard<any, any> {
  override styling(): string {
    return "styling-class-name";
  }

  override layout(): JSX.Element | undefined {
    return (
      <>
        <ListWidget />
        <ChartWidget />
      </>
    );
  }
}
```

## React Hook list

### useData
Fundamental helper hook function to do asynchronized operation like fetch data from a remote database / backend API.
It returns custom data, loading status, error object and reload function.
By default, it will fetch the data once the component has been initialized.

### useTeams
The hook is based on the Microsoft Teams JavaScript SDK, the Fluent UI components and Microsoft Graph Toolkit, merged from [msteams-react-base-component](https://github.com/wictorwilen/msteams-react-base-component).
It returns a tuple of where an object of properties are in the first field and an object of methods in the second.
If you want to manually set the initial theme, please pass the config object to `useTeams()`.

### useTeamsFx
Initialize TeamsFx and Teams JS SDK, in a development environment, verbose logging message will be printed to console.
It returns the TeamsFx instance as data.
If you want to customize TeamsFx like customizing setting, please pass the config object to `useTeamsFx()`.
> Note: `useTeamsFx()` has been deprecated and will be removed in the future release. Please use `useTeamsUserCredential()` instead as below:

```ts
const authConfig: TeamsUserCredentialAuthConfig = {
  clientId: process.env.REACT_APP_CLIENT_ID,
  initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
}
const { loading, theme, themeString, teamsUserCredential } = useTeamsUserCredential(authConfig);
```

### useGraph
This hook function leverage `useData` to call Graph API. It will execute the fetchGraphDataAsync function that the developer passes in first.
If user has not consented to the scopes of AAD resources, `useGraph()`, `useGraphWithCredential` will automatically call login function to pop up the consent dialog. So, developers can focus on the business logic of how to fetch Microsoft Graph data.

## Dashboard Related Classes

### BaseDashboard
The BaseDashboard is a React component that provides a basic dashboard layout implementation for developers to quickly build a dashboard tab for Microsoft Teams. You can inherit this class and override some methods to customize your own dashboard. For example, define the layout of the widget in your dashboard by overriding the `layout()` method, and customize the dashboard style by overriding the `styling()` method.

### BaseWidget
The BaseWidget is a React component that provides a basic widget layout implementation for developers to quickly build a widget. You can inherit this class and override some methods to customize your own widget. For example, define the header of the widget by overriding the `header()` method, and get data needed for the widget by overriding the `getData()` method, etc.

## Next steps

Please take a look at the [Samples](https://github.com/OfficeDev/TeamsFx-Samples) project for detailed examples on how to use this library.

## Related projects

- [Microsoft Teams Toolkit for Visual Studio Code](https://github.com/OfficeDev/TeamsFx/tree/main/packages/vscode-extension)
- [TeamsFx SDK](https://github.com/OfficeDev/TeamsFx/tree/main/packages/sdk)

## Data Collection.

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at https://go.microsoft.com/fwlink/?LinkID=824704. You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)

If you are interested in fixing issues and contributing directly to the code base, please see the [Contributing Guide](./CONTRIBUTING.md).

## Reporting Security Issues

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them to the Microsoft Security Response Center (MSRC) at [https://msrc.microsoft.com/create-report](https://msrc.microsoft.com/create-report).

If you prefer to submit without logging in, send email to [secure@microsoft.com](mailto:secure@microsoft.com). If possible, encrypt your message with our PGP key; please download it from the the [Microsoft Security Response Center PGP Key page](https://www.microsoft.com/en-us/msrc/pgp-key-msrc).

You should receive a response within 24 hours. If for some reason you do not, please follow up via email to ensure we received your original message. Additional information can be found at [microsoft.com/msrc](https://www.microsoft.com/msrc).

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE.txt) license.
