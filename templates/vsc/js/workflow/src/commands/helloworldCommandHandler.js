const helloWorldCard = require("../adaptiveCards/helloworldCommandResponse.json");
const ACData = require("adaptivecards-templating");
const { CardFactory, MessageFactory } = require("botbuilder");

class HelloWorldCommandHandler {
  triggerPatterns = "helloWorld";

  async handleCommandReceived(context, state) {
    console.log(`Bot received message: ${context.activity.text}`);

    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "Your Hello World Bot is Running",
        body: "Congratulations! Your hello world bot is running. Click the button below to trigger an action.",
      },
    });
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}

module.exports = {
  HelloWorldCommandHandler,
};
