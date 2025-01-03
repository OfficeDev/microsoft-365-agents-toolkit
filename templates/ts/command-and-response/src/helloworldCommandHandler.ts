import { Selector } from "@microsoft/teams-ai";
import * as ACData from "adaptivecards-templating";
import { Activity, CardFactory, MessageFactory, TurnContext } from "botbuilder";
import helloWorldCard from "./adaptiveCards/helloworldCommand.json";
import { ApplicationTurnState } from "./internal/interface";

/**
 * The `HelloWorldCommandHandler` registers a pattern and responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
export class HelloWorldCommandHandler {
  triggerPatterns: string | RegExp | Selector | (string | RegExp | Selector)[] = "helloWorld";

  async handleCommandReceived(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<string | Partial<Activity> | void> {
    console.log(`App received message: ${context.activity.text}`);

    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "Your Hello World App is Running",
        body: "Congratulations! Your Hello World App is running. Open the documentation below to learn more about how to build applications with the Teams Toolkit.",
      },
    });
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}
