import "./components/github.scss";
import "./components/document.scss";

import * as React from "react";
import { useEffect } from "react";
import { TelemetryTriggerFrom } from "../../telemetry/extTelemetryEvents";
import CollapsibleStep from "./components/collapsibleStep";
import ExternalLink from "./components/externalLink";
import M365Account from "../../../img/webview/accountHelp/ttk-m365-account.png";
import ButtonLink from "./components/buttonLink";

export default function PrepareM365Account() {
  const [theme, setTheme] = React.useState("light");

  useEffect(() => {
    let currentTheme = document.body.className;
    const prefix = "vscode-";
    if (currentTheme.startsWith(prefix)) {
      // strip prefix
      currentTheme = currentTheme.substring(prefix.length);
    }

    if (currentTheme === "high-contrast") {
      currentTheme = "dark"; // the high-contrast theme seems to be an extreme case of the dark theme
    }

    if (theme === currentTheme) return;
    setTheme(currentTheme);
  }, []);

  return (
    <div className="markdown-body">
      <h1 id="prepare-a-qualified-Microsoft-365-account-for-Teams-app-development">
        Prepare a Qualified Microsoft 365 Account for Teams App Development
      </h1>
      <h2 id="account-requirements">Account Requirements </h2>
      <p>
        The following two conditions are required for Teams app development:
        <ol>
          <li>
            The Microsoft 365 account should be your work or school account, not your personal
            account.
          </li>
          <li>The Microsoft 365 account needs permission to upload custom app in Teams.</li>
        </ol>
      </p>
      <p>
        You can contact your tenant administrator to turn on the upload custom app permission for
        your organization.{" "}
      </p>
      <p>
        Or, if you're a Visual Studio subscriber, create a Microsoft 365 developer account to
        resolve your account issues.{" "}
        <ExternalLink
          title="Get more info about Microsoft 365 Developer Program"
          link="https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program"
          triggerFrom={TelemetryTriggerFrom.AccountHelp}
        />
      </p>

      <h2 id="how">How</h2>
      <div className="collapsibleSteps">
        <CollapsibleStep
          step={1}
          title="Create Microsoft 365 Development Account"
          triggerFrom={TelemetryTriggerFrom.AccountHelp}
          identifier="account-help-step1"
        >
          <p>
            Select the button below to create an instant sandbox and get your developer account.
          </p>
          <p>
            <ButtonLink
              title="Sign up for Microsoft 365 developer program for free"
              link="https://developer.microsoft.com/en-us/microsoft-365/dev-program"
              triggerFrom={TelemetryTriggerFrom.AccountHelp}
            />
          </p>
          <p>
            For more information, visit the{" "}
            <ExternalLink
              title="Set up a developer subscription documentation"
              link="https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-get-started"
              triggerFrom={TelemetryTriggerFrom.AccountHelp}
            />
          </p>
          <p>
            <video width={800} height={450} controls src="https://aka.ms/teamsfx-ITP-video" />
          </p>
          <blockquote>
            <p>
              Use the Administrator account (*.onmicrosoft.com) email address to log in to your
              development environment.
            </p>
          </blockquote>
        </CollapsibleStep>
      </div>
      <div className="collapsibleSteps">
        <CollapsibleStep
          step={2}
          title="Use your development account in Teams Toolkit for Visual Studio Code"
          triggerFrom={TelemetryTriggerFrom.AccountHelp}
          identifier="account-help-step2"
        >
          <p>
            Open Teams Toolkit for Visual Studio Code and log into the Teams Toolkit extension using
            your developer account created in step 1.
          </p>
          <p>The permission to upload custom app is already set up.</p>
          <p>
            <img src={M365Account} alt="Teams Toolkit Microsoft 365 account" />
          </p>
        </CollapsibleStep>
      </div>
    </div>
  );
}
