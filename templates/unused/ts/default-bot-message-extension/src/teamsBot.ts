import { default as axios } from "axios";
import * as querystring from "querystring";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MessagingExtensionAction,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  MessagingExtensionActionResponse,
  AppBasedLinkQuery,
} from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json";
import rawLearnCard from "./adaptiveCards/learn.json";
import searchResultCard from "./adaptiveCards/searchResultCard.json";
import actionCard from "./adaptiveCards/actionCard.json";
import linkUnfurlingCard from "./adaptiveCards/linkUnfurlingCard.json";
import * as ACData from "adaptivecards-templating";

export interface DataInterface {
  likeCount: number;
}

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };

  constructor() {
    super();

    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const welcomeAttachement = CardFactory.adaptiveCard(rawWelcomeCard);
          await context.sendActivity({ attachments: [welcomeAttachement] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const template = new ACData.Template(rawLearnCard);
          const learnCard = template.expand({ $root: this.likeCountObj });
          const learnAttachment = CardFactory.adaptiveCard(learnCard);
          await context.sendActivity({ attachments: [learnAttachment] });
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const welcomeAttachement = CardFactory.adaptiveCard(rawWelcomeCard);
          await context.sendActivity({ attachments: [welcomeAttachement] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const template = new ACData.Template(rawLearnCard);
      const learnCard = template.expand({ $root: this.likeCountObj });
      const learnAttachment = CardFactory.adaptiveCard(learnCard);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [learnAttachment],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    }
  }

  // Message extension Code
  // Action.
  public async handleTeamsMessagingExtensionSubmitAction(
    context: TurnContext,
    action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
    const template = new ACData.Template(actionCard);
    const card = template.expand({
      $root: {
        title: action.data.title ?? "",
        subTitle: action.data.subTitle ?? "",
        text: action.data.text ?? "",
      },
    });
    const attachment = CardFactory.adaptiveCard(card);
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
      },
    };
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const template = new ACData.Template(searchResultCard);
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

  // Link Unfurling.
  public async handleTeamsAppBasedLinkQuery(
    context: TurnContext,
    query: AppBasedLinkQuery
  ): Promise<MessagingExtensionResponse> {
    // When the returned card is an adaptive card, the previewCard property of the attachment is required.
    const previewCard = CardFactory.thumbnailCard("Preview Card", query.url, [
      "https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png",
    ]);

    const attachment = { ...CardFactory.adaptiveCard(linkUnfurlingCard), preview: previewCard };

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
        suggestedActions: {
          actions: [
            {
              title: "default",
              type: "setCachePolicy",
              value: '{"type":"no-cache"}',
            },
          ],
        },
      },
    };
  }
}
