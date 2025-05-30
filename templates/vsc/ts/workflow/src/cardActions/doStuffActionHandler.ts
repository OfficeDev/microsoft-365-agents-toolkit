import { AdaptiveCard } from "@microsoft/teams-ai";
import * as ACData from "adaptivecards-templating";
import { TurnContext } from "botbuilder";
import responseCard from "../adaptiveCards/doStuffActionResponse.json";

/**
 * The `DoStuffActionHandler` responds
 * with an Adaptive Card if the user clicks the Adaptive Card action with `triggerVerb`.
 */
export class DoStuffActionHandler {
  /**
   * A global unique string associated with the `Action.Execute` action.
   * The value should be the same as the `verb` property which you define in your adaptive card JSON.
   */
  triggerVerb = "doStuff";

  async handleActionInvoked(context: TurnContext, actionData: any): Promise<string | AdaptiveCard> {
    /**
     * You can send an adaptive card to respond to the card action invoke.
     */
    const cardJson = new ACData.Template(responseCard).expand({
      $root: {
        title: "Hello World Bot",
        body: "Congratulations! Your task is processed successfully.",
      },
    });
    return cardJson as AdaptiveCard;

    /**
     * If you want to send invoke response with text message, you can:
     * 
     return "[ACK] Successfully!";
    */
  }
}
