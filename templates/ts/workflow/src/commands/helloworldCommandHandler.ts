import { Selector, CardFactory, MessageFactory, TurnContext } from "@microsoft/agents-hosting";
import { Activity } from "@microsoft/agents-activity";
import * as ACData from "adaptivecards-templating";
import helloWorldCard from "../adaptiveCards/helloworldCommandResponse.json";
import { ApplicationTurnState } from "../internal/interface";

/**
 * The `HelloWorldCommandHandler` responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
export class HelloWorldCommandHandler {
  triggerPatterns: string | RegExp | Selector | (string | RegExp | Selector)[] = "helloWorld";

  async handleCommandReceived(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<string | Activity | void> {
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
