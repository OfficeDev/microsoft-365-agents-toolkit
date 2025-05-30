const ACData = require("adaptivecards-templating");
const responseCard = require("../adaptiveCards/doStuffActionResponse.json");

class DoStuffActionHandler {
  /**
   * A global unique string associated with the `Action.Execute` action.
   * The value should be the same as the `verb` property which you define in your adaptive card JSON.
   */
  triggerVerb = "doStuff";

  async handleActionInvoked(context, actionData) {
    /**
     * You can send an adaptive card to respond to the card action invoke.
     */
    const cardJson = new ACData.Template(responseCard).expand({
      $root: {
        title: "Hello World Bot",
        body: "Congratulations! Your task is processed successfully.",
      },
    });
    return cardJson;

    /**
     * If you want to send invoke response with text message, you can:
     * 
     return "[ACK] Successfully!";
     */
  }
}

module.exports = {
  DoStuffActionHandler,
};
