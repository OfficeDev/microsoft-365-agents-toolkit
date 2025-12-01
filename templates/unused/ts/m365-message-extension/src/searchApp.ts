import { default as axios } from "axios";
import * as querystring from "querystring";
import { CardFactory, TurnContext } from "@microsoft/agents-hosting";
import {
  TeamsActivityHandler,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
} from "@microsoft/agents-hosting-teams";
import * as ACData from "adaptivecards-templating";
import helloWorldCard from "./adaptiveCards/helloWorldCard.json";

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;

    // Due to npmjs search limitations, do not search if input length < 2
    if (searchQuery.length < 2) {
      return {
        composeExtension: {
          type: "result",
          attachmentLayout: "list",
          attachments: [],
        },
      };
    }

    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const template = new ACData.Template(helloWorldCard);
      const card = template.expand({
        $root: {
          name: obj.package.name,
          description: obj.package.description,
        },
      });
      const preview = CardFactory.heroCard(obj.package.name);
      const attachment = { ...CardFactory.adaptiveCard(card), preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }
}
