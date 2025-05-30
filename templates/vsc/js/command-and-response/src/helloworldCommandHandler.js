const helloWorldCard = require("./adaptiveCards/helloworldCommand.json");
const { CardFactory, MessageFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");

class HelloWorldCommandHandler {
  triggerPatterns = "helloWorld";

  async handleCommandReceived(context, state) {
    // verify the command arguments which are received from the client if needed.
    console.log(`App received message: ${context.activity.text}`);

    // do something to process your command and return message activity as the response
    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "Your Hello World App is Running",
        body: "Congratulations! Your Hello World App is running. Open the documentation below to learn more about how to build applications with the Microsoft 365 Agents Toolkit.",
      },
    });
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}

module.exports = {
  HelloWorldCommandHandler,
};
