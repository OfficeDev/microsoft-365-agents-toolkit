/**
 * The `GenericCommandHandler` responds
 * with appropriate messages if the user types general command inputs, such as "hi", "hello", and "help".
 */
export class GenericCommandHandler {
  async handleCommandReceived(text: string): Promise<string> {
    let response = "";
    switch (text) {
      case "hi":
        response =
          "Hi there! I'm your Workflow Bot, here to assist you with your tasks. Type 'help' for a list of available commands.";
        break;
      case "hello":
        response =
          "Hello! I'm your Workflow Bot, always ready to help you out. If you need assistance, just type 'help' to see the available commands.";
        break;
      case "help":
        response =
          "Here's a list of commands I can help you with:\n" +
          "- 'hi' or 'hello': Say hi or hello to me, and I'll greet you back.\n" +
          "- 'help': Get a list of available commands.\n" +
          "- 'helloWorld': See a sample workflow from me.\n" +
          "\nFeel free to ask for help anytime you need it!";
        break;
      default:
        response = `Sorry, command unknown. Please type 'help' to see the list of available commands.`;
    }

    return response;
  }
}
