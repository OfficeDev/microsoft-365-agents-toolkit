import * as ACData from "adaptivecards-templating";
import helloWorldCard from "../adaptiveCards/helloworldCommandResponse.json";

/**
 * The `HelloWorldCommandHandler` responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
export class HelloWorldCommandHandler {
  triggerPatterns = "helloWorld";

  shouldTrigger(text: string | undefined): boolean {
    return text === this.triggerPatterns;
  }

  async handleCommandReceived(text: string): Promise<any> {
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
