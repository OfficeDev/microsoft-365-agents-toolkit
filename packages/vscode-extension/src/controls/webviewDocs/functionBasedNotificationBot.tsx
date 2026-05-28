import "./components/github.scss";
import "./components/document.scss";

import * as React from "react";

import {
  InProductGuideInteraction,
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../telemetry/extTelemetryEvents";
import { Commands } from "../Commands";
import CollapsibleSection from "./components/collapsibleSection";
import CollapsibleStep from "./components/collapsibleStep";
import ExternalLink from "./components/externalLink";
import { useEffect } from "react";
import NotificationAdaptiveCard from "../../../img/webview/guide/notification-sends-adaptive-card.png";

export default function FunctionBasedNotificationBot() {
  let scrollToBottom = false;

  const [theme, setTheme] = React.useState("light");
  const name = "function-based-notification-bot";

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

    const handleScroll = () => {
      console.log(window.scrollY);
      if (!scrollToBottom && window.scrollY > 1000) {
        scrollToBottom = true;
        vscode.postMessage({
          command: Commands.SendTelemetryEvent,
          data: {
            eventName: TelemetryEvent.InteractWithInProductDoc,
            properties: {
              [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.InProductDoc,
              [TelemetryProperty.Interaction]: InProductGuideInteraction.ScrollToBottom,
              [TelemetryProperty.TutorialName]: name,
            },
          },
        });
      }
    };

    window.addEventListener("scroll", handleScroll);

    return () => {
      window.removeEventListener("scroll", handleScroll);
    };
  }, []);

  return (
    <div className="markdown-body">
      <h1 id="overview-of-the-notification-bot-template">
        Overview of the Notification bot template
      </h1>
      <p>
        This template showcases an app that send a message to Teams with Adaptive Cards triggered by
        a HTTP post request or timer schedule. You can further extend the template to consume,
        transform and post events to individual, chat or channel in Teams.
      </p>
      <p>
        The app template is built using the TeamsFx SDK, which provides a simple set of functions
        over the Microsoft Bot Framework to implement this scenario.
      </p>
      <h2 id="get-started-with-the-notification-bot">Get Started with the Notification bot</h2>
      <blockquote>
        <p>
          <strong>Prerequisites</strong>
        </p>
        <p>To run the notification bot template in your local dev machine, you will need:</p>
        <ul>
          <li>
            <ExternalLink
              title="Node.js"
              link="https://nodejs.org/"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />
            , supported versions: 14, 16, 18 (preview)
          </li>
          <li>
            An{" "}
            <ExternalLink
              title="Microsoft 365 account for development"
              link="https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />
          </li>
        </ul>
        <p>
          <strong>Note</strong>
        </p>
        <p>
          Your app can be installed into a team, or a group chat, or as personal app. See{" "}
          <ExternalLink
            title="Installation and Uninstallation"
            link="https://aka.ms/teamsfx-notification#customize-installation"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
          .
        </p>
      </blockquote>
      <ol>
        <li>First, select the Teams Toolkit icon on the left in the VS Code toolbar.</li>
        <li>
          In the Account section, sign in with your{" "}
          <ExternalLink
            title="Microsoft 365 account"
            link="https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />{" "}
          if you haven&#39;t already.
        </li>
        <li>
          Press F5 to start debugging which launches your app in Teams using a web browser. Select{" "}
          <code className="code">Debug in Teams (Edge)</code> or{" "}
          <code className="code">Debug in Teams (Chrome)</code>.
        </li>
        <li>
          When Teams launches in the browser, select the Add button in the dialog to install your
          app to Teams.
        </li>
        <li>
          <p>
            If you select <code className="code">Timer Trigger</code>, wait for 30 seconds. If you
            select <code className="code">HTTP Trigger</code>, send a POST request to{" "}
            <code className="code">http://&lt;endpoint&gt;/api/notification</code> with your
            favorite tool (like <code className="code">Postman</code>)
          </p>
          <ul>
            <li>
              When your project is running locally, replace{" "}
              <code className="code">&lt;endpoint&gt;</code> with{" "}
              <code className="code">localhost:3978</code>
            </li>
            <li>
              When your project is deployed to Azure App Service, replace{" "}
              <code className="code">&lt;endpoint&gt;</code> with the url from Azure App Service
            </li>
          </ul>
        </li>
      </ol>
      <p>The bot will send an Adaptive Card to Teams:</p>
      <p>
        <img src={NotificationAdaptiveCard} alt="Notification Message in Teams" />
      </p>
      <CollapsibleSection
        title="What's included in the template"
        triggerFrom={TelemetryTriggerFrom.InProductDoc}
        identifier="function-based-notification-bot-template-contents"
      >
        <table>
          <thead>
            <tr>
              <th>Folder</th>
              <th>Contents</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                <code className="code">.fx</code>
              </td>
              <td>Project level settings, configurations, and environment information</td>
            </tr>
            <tr>
              <td>
                <code className="code">.vscode</code>
              </td>
              <td>VSCode files for local debug</td>
            </tr>
            <tr>
              <td>
                <code className="code">bot</code>
              </td>
              <td>The source code for the notification Teams application</td>
            </tr>
            <tr>
              <td>
                <code className="code">templates</code>
              </td>
              <td>
                Templates for the Teams application manifest and for provisioning Azure resources
              </td>
            </tr>
          </tbody>
        </table>
        <p>
          The following files can be customized and demonstrate an example implementation to get you
          started.
        </p>
        <table>
          <thead>
            <tr>
              <th>File</th>
              <th>Contents</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                <code className="code">*Trigger/function.json</code>
              </td>
              <td>Azure Function bindings for the notification trigger</td>
            </tr>
            <tr>
              <td>
                <code className="code">src/*Trigger.js</code>
              </td>
              <td>Notification trigger implementation</td>
            </tr>
            <tr>
              <td>
                <code className="code">src/teamsBot.js</code>
              </td>
              <td>An empty teams activity handler for bot customization</td>
            </tr>
            <tr>
              <td>
                <code className="code">src/adaptiveCards/notification-default.json</code>
              </td>
              <td>A generated Adaptive Card that is sent to Teams</td>
            </tr>
          </tbody>
        </table>
        <p>
          The following files implement the core notification on the Bot Framework. You generally
          will not need to customize these files.
        </p>
        <table>
          <thead>
            <tr>
              <th>File / Folder</th>
              <th>Contents</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                <code className="code">src/internal/initialize.js</code>
              </td>
              <td>Application initialization</td>
            </tr>
            <tr>
              <td>
                <code className="code">messageHandler/</code>
              </td>
              <td>Azure Function bindings to implement Bot protocol</td>
            </tr>
            <tr>
              <td>
                <code className="code">src/internal/messageHandler.js</code>
                <br />
                <code className="code">src/internal/responseWrapper.js</code>
              </td>
              <td>Bot protocol implementation</td>
            </tr>
          </tbody>
        </table>
        <p>
          The following files are project-related files. You generally will not need to customize
          these files.
        </p>
        <table>
          <thead>
            <tr>
              <th>File / Folder</th>
              <th>Contents</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                <code className="code">.funcignore</code>
              </td>
              <td>Azure Functions ignore file to exclude local files</td>
            </tr>
            <tr>
              <td>
                <code className="code">.gitignore</code>
              </td>
              <td>Git ignore file</td>
            </tr>
            <tr>
              <td>
                <code className="code">host.json</code>
              </td>
              <td>Azure Functions host file</td>
            </tr>
            <tr>
              <td>
                <code className="code">local.settings.json</code>
              </td>
              <td>Azure Functions settings for local debugging</td>
            </tr>
            <tr>
              <td>
                <code className="code">package.json</code>
              </td>
              <td>NPM package file</td>
            </tr>
          </tbody>
        </table>
      </CollapsibleSection>
      <h2 id="extend-the-notification-bot-template">Extend the notification bot template</h2>
      <p>
        There are few customizations you can make to extend the template to fit your business
        requirements.
      </p>
      <div className="collapsibleSteps">
        <CollapsibleStep
          step={1}
          title="Customize the trigger point from event source"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          identifier="function-based-notification-bot-step1"
        >
          <p>
            If you selected <code className="code">timer</code> trigger, the default Azure Function
            timer trigger (<code className="code">src/timerTrigger.js</code>) implementation simply
            sends a hard-coded Adaptive Card every 30 seconds. You can edit the file
            <code className="code">*Trigger/function.json</code> to customize the
            <code className="code">schedule</code> property. Refer to the{" "}
            <ExternalLink
              title="Azure Function documentation"
              link="https://docs.microsoft.com/azure/azure-functions/functions-bindings-timer?tabs=in-process&amp;pivots=programming-language-javascript#ncrontab-expressions"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />
            for more details.
          </p>
          <p>
            If you selected <code className="code">http</code> trigger, when this trigger is hit
            (via a HTTP request), the default implementation sends a hard-coded Adaptive Card to
            Teams. You can change this behavior by customizing
            <code className="code">src/*Trigger.js</code>. A typical implementation might make an
            API call to retrieve some events and/or data, and then send an Adaptive Card as
            appropriate.
          </p>
          <p>You can also add any Azure Function trigger. For example:</p>
          <ul>
            <li>
              You can use an <code className="code">Event Hub</code> trigger to send notifications
              when an event is pushed to Azure Event Hub.
            </li>
            <li>
              You can use a <code className="code">Cosmos DB</code> trigger to send notifications
              when a Cosmos document has been created or updated.
            </li>
          </ul>
          <p>
            See Azure Functions{" "}
            <ExternalLink
              title="supported triggers"
              link="https://docs.microsoft.com/azure/azure-functions/functions-triggers-bindings?tabs=javascript#supported-bindings"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />
            .
          </p>
        </CollapsibleStep>
      </div>
      <div className="collapsibleSteps">
        <CollapsibleStep
          step={2}
          title="Customize the notification content"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          identifier="function-based-notification-bot-step2"
        >
          <p>
            <code className="code">src/adaptiveCards/notification-default.json</code> defines the
            default Adaptive Card. You can use the
            <ExternalLink
              title="Adaptive Card Designer"
              link="https://adaptivecards.io/designer/"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />
            to help visually design your Adaptive Card UI.
          </p>
          <p>
            <code className="code">src/cardModels.js</code> defines a data structure that is used to
            fill data for the Adaptive Card. The binding between the model and the Adaptive Card is
            done by name matching (for example,<code className="code">CardData.title</code>
            maps to <code className="code">{"${title}"}</code> in the Adaptive Card). You can add,
            edit, edit, or remove properties and their bindings to customize the Adaptive Card to
            your needs.
          </p>
          <p>
            You can also add new cards if needed. Follow this
            <ExternalLink
              title="sample"
              link="https://aka.ms/teamsfx-adaptive-card-sample"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />{" "}
            to see how to build different types of adaptive cards with a list or a table of dynamic
            contents using <code className="code">ColumnSet</code> and{" "}
            <code className="code">FactSet</code>.
          </p>
        </CollapsibleStep>
      </div>
      <div className="collapsibleSteps">
        <CollapsibleStep
          step={3}
          title="Customize where notifications are sent"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          identifier="function-based-notification-bot-step3"
        >
          <p>Notifications can be sent to where the bot is installed:</p>
          <ul>
            <li>
              <ExternalLink
                title="Send notifications to a channel"
                link="https://aka.ms/teamsfx-notification#send-notifications-to-a-channel"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Send notifications to a group chat"
                link="https://aka.ms/teamsfx-notification#send-notifications-to-a-group-chat"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Send notifications to a personal chat"
                link="https://aka.ms/teamsfx-notification#send-notifications-to-a-personal-chat"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
          </ul>
          <p>You can also send the notifications to a specific receiver:</p>
          <ul>
            <li>
              <ExternalLink
                title="Send notifications to a specific channel"
                link="https://aka.ms/teamsfx-notification#send-notifications-to-a-specific-channel"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Send notifications to a specific person"
                link="https://aka.ms/teamsfx-notification#send-notifications-to-a-specific-person"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
          </ul>
          <p>
            Congratulations, you&#39;ve just created your own notification! To get more info about
            extending the notification bot template,{" "}
            <ExternalLink
              title="visit the documentation on GitHub"
              link="https://aka.ms/teamsfx-notification"
              triggerFrom={TelemetryTriggerFrom.InProductDoc}
              docName={name}
            />
            . You can find more scenarios like:
          </p>
          <ul>
            <li>
              <ExternalLink
                title="Customize storage"
                link="https://aka.ms/teamsfx-notification#customize-storage"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Customize adapter"
                link="https://aka.ms/teamsfx-notification#customize-adapter"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Customize the way to initialize the bot"
                link="https://aka.ms/teamsfx-notification#customize-initialization"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Add authentication for your notification API"
                link="https://aka.ms/teamsfx-notification#add-authentication-for-your-notification-api"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Connect to existing APIs"
                link="https://aka.ms/teamsfx-notification#connect-to-existing-api"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
            <li>
              <ExternalLink
                title="Frequently asked questions"
                link="https://aka.ms/teamsfx-notification#frequently-asked-questions"
                triggerFrom={TelemetryTriggerFrom.InProductDoc}
                docName={name}
              />
            </li>
          </ul>
        </CollapsibleStep>
      </div>
      <h2 id="extend-notification-bot-with-other-bot-scenarios">
        Extend notification bot with other bot scenarios
      </h2>
      <p>
        Notification bot is compatible with other bot scenarios like command bot and workflow bot.
      </p>
      <h3 id="add-command-to-your-application">Add command to your application</h3>
      <p>
        The command and response feature adds the ability for your application to &quot;listen&quot;
        to commands sent to it via a Teams message and respond to commands with Adaptive Cards.
        Follow the{" "}
        <ExternalLink
          title="steps here"
          link="https://aka.ms/teamsfx-command-response#How-to-add-more-command-and-response"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          docName={name}
        />{" "}
        to add the command response feature to your notification bot. Refer{" "}
        <ExternalLink
          title="the command bot document"
          link="https://aka.ms/teamsfx-command-response"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          docName={name}
        />{" "}
        for more information.
      </p>
      <h3 id="add-workflow-to-your-notification-bot">Add workflow to your notification bot</h3>
      <p>
        Adaptive cards can be updated on user action to allow user progress through a series of
        cards that require user input. Developers can define actions and use a bot to return an
        Adaptive Cards in response to user action. This can be chained into sequential workflows.
        Follow the{" "}
        <ExternalLink
          title="steps here"
          link="https://aka.ms/teamsfx-card-action-response#add-more-card-actions"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          docName={name}
        />{" "}
        to add workflow feature to your notification bot. Refer{" "}
        <ExternalLink
          title="the workflow document"
          link="https://aka.ms/teamsfx-card-action-response"
          triggerFrom={TelemetryTriggerFrom.InProductDoc}
          docName={name}
        />{" "}
        for more information.
      </p>
      <h2 id="additional-information-and-references">Additional information and references</h2>
      <ul>
        <li>
          <ExternalLink
            title="Manage multiple environments"
            link="https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
        </li>
        <li>
          <ExternalLink
            title="Collaborate with others"
            link="https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
        </li>
        <li>
          <ExternalLink
            title="Teams Toolkit Documentations"
            link="https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
        </li>
        <li>
          <ExternalLink
            title="Teams Toolkit CLI"
            link="https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-cli"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
        </li>
        <li>
          <ExternalLink
            title="TeamsFx SDK"
            link="https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
        </li>
        <li>
          <ExternalLink
            title="Teams Toolkit Samples"
            link="https://github.com/OfficeDev/TeamsFx-Samples"
            triggerFrom={TelemetryTriggerFrom.InProductDoc}
            docName={name}
          />
        </li>
      </ul>
    </div>
  );
}
