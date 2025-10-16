const ACData = require("adaptivecards-templating");
const helloWorldCard = require("./adaptiveCards/helloworldCommand.json");

class HelloWorldCommandHandler {
  constructor() {
    this.triggerPattern = "helloWorld";
  }

  canHandle(text) {
    return text === this.triggerPattern;
  }

  async handleCommandReceived(text) {
    console.log(`App received message: ${text}`);

    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "Your Hello World App is Running",
        body: "Congratulations! Your Hello World App is running. Open the documentation below to learn more about how to build applications with the Microsoft 365 Agents Toolkit.",
      },
    });

    return cardJson;
  }
}

module.exports = {
  HelloWorldCommandHandler,
};
