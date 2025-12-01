const { CardFactory } = require("@microsoft/agents-hosting");
const { TeamsActivityHandler } = require("@microsoft/agents-hosting-teams");
const ACData = require("adaptivecards-templating");
const helloWorldCard = require("./adaptiveCards/helloWorldCard.json");

class ActionApp extends TeamsActivityHandler {
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const template = new ACData.Template(helloWorldCard);
    const card = template.expand({
      $root: {
        title: action.data.title ?? "",
        subTitle: action.data.subTitle ?? "",
        text: action.data.text ?? "",
      },
    });
    const attachment = CardFactory.adaptiveCard(card);
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
      },
    };
  }
}
module.exports.ActionApp = ActionApp;
