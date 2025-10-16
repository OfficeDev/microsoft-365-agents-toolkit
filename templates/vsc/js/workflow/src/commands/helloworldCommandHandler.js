const ACData = require("adaptivecards-templating");
const helloWorldCard = require("../adaptiveCards/helloworldCommandResponse.json");

/**
 * The `HelloWorldCommandHandler` responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
class HelloWorldCommandHandler {
  constructor() {
    this.triggerPatterns = "helloWorld";
  }

  shouldTrigger(text) {
    return text === this.triggerPatterns;
  }

  async handleCommandReceived(text) {
    console.log(`Bot received message: ${text}`);

    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "Your Hello World Bot is Running",
        body: "Congratulations! Your hello world bot is running. Click the button below to trigger an action.",
      },
    });

    return cardJson;
  }
}

module.exports = { HelloWorldCommandHandler };
