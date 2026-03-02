import * as ACData from "adaptivecards-templating";
import helloWorldCard from "./adaptiveCards/helloworldCommand.json";

/**
 * The `HelloWorldCommandHandler` registers a pattern and responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
export class HelloWorldCommandHandler {
  private triggerPattern = "helloWorld";

  canHandle(text: string): boolean {
    return text === this.triggerPattern;
  }

  async handleCommandReceived(text: string): Promise<any> {
    console.log(`App received message: ${text}`);

    // Use ACData templating to expand the card with data
    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "Your Hello World App is Running",
        body: "Congratulations! Your Hello World App is running. Open the documentation below to learn more about how to build applications with the Microsoft 365 Agents Toolkit.",
      },
    });

    // Return the expanded adaptive card data
    return cardJson;
  }
}
