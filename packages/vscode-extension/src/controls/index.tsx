import * as React from "react";
import * as ReactDOM from "react-dom";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { IntlProvider } from "react-intl";
import { MemoryRouter, Route } from "react-router-dom";

import { PanelType } from "./PanelType";
import SampleGallery from "./sampleGallery/SampleGallery";
import AccountHelp from "./webviewDocs/accountHelp";
import FunctionBasedNotificationBot from "./webviewDocs/functionBasedNotificationBot";
import ExpressServerNotificationBot from "./webviewDocs/expressServerNotificationBot";
import WorkflowBot from "./webviewDocs/workflowBot";

const language = "en";

function getBodyTheme() {
  const themeClass = document.body.className;
  return themeClass.includes("dark") || themeClass.includes("high-contrast")
    ? webDarkTheme
    : webLightTheme;
}

function ThemedApp() {
  const [theme, setTheme] = React.useState(getBodyTheme);

  React.useEffect(() => {
    const observer = new MutationObserver(() => {
      setTheme(getBodyTheme());
    });

    observer.observe(document.body, {
      attributes: true,
      attributeFilter: ["class"],
    });

    return () => observer.disconnect();
  }, []);

  return (
    <FluentProvider theme={theme}>
      <App />
    </FluentProvider>
  );
}

ReactDOM.render(
  <IntlProvider locale={language}>
    <ThemedApp />
  </IntlProvider>,
  document.getElementById("root") as HTMLElement
);

function App(props: any) {
  let initialIndex = 0;
  if (panelType === PanelType.RespondToCardActions) {
    initialIndex = 1;
  } else if (panelType === PanelType.AccountHelp) {
    initialIndex = 2;
  } else if (panelType === PanelType.FunctionBasedNotificationBotReadme) {
    initialIndex = 3;
  } else if (panelType === PanelType.ExpressServerNotificationBotReadme) {
    initialIndex = 4;
  }

  return (
    <MemoryRouter
      initialEntries={[
        "/sample-gallery",
        "/respond-to-card-actions",
        "/account-help",
        "/function-based-notification-bot",
        "/express-server-notification-bot",
      ]}
      initialIndex={initialIndex}
    >
      <Route
        path="/sample-gallery"
        render={() => (
          <SampleGallery
            shouldShowChat={shouldShowChat}
            shouldHideTeamsAgentPreviewTag={shouldHideTeamsAgentPreviewTag}
          />
        )}
      />
      <Route path="/respond-to-card-actions" component={WorkflowBot} />
      <Route path="/account-help" component={AccountHelp} />
      <Route path="/function-based-notification-bot" component={FunctionBasedNotificationBot} />
      <Route path="/express-server-notification-bot" component={ExpressServerNotificationBot} />
    </MemoryRouter>
  );
}
