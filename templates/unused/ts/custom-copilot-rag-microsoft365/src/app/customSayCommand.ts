import { ActivityTypes, Channels, TurnContext } from "botbuilder";
import { PredictedSayCommand, TurnState, Utilities } from "@microsoft/teams-ai";
import { AIEntity, ClientCitation } from "@microsoft/teams-ai/lib/types";

export function sayCommand<TState extends TurnState = TurnState>(feedbackLoopEnabled = false) {
  return async (context: TurnContext, _state: TState, data: PredictedSayCommand) => {
    if (!data.response?.content) {
      return "";
    }

    const isTeamsChannel = context.activity.channelId === Channels.Msteams;
    let content = "";
    let result = undefined;
    try {
      result = JSON.parse(data.response.content);
    } catch (error) {
      console.error(`Response is not valid json, send the raw text. error: ${error}`);
      await context.sendActivity({
        type: ActivityTypes.Message,
        text: data.response.content,
        ...(isTeamsChannel ? { channelData: { feedbackLoopEnabled } } : {}),
        entities: [
          {
            type: "https://schema.org/Message",
            "@type": "Message",
            "@context": "https://schema.org",
            "@id": "",
            additionalType: ["AIGeneratedContent"],
          },
        ] as AIEntity[],
      });
      return "";
    }

    // If the response from AI includes citations, those citations will be parsed and added to the SAY command.
    const citations = [];
    let position = 1;

    if (result.results && result.results.length > 0) {
      result.results.forEach((contentItem) => {
        if (contentItem.citationTitle && contentItem.citationTitle.length > 0) {
          const clientCitation: ClientCitation = {
            "@type": "Claim",
            position: position,
            appearance: {
              "@type": "DigitalDocument",
              name: contentItem.citationTitle || `Document #${position}`,
              url: contentItem.citationUrl,
              abstract: Utilities.snippet(contentItem.citationContent, 500),
            },
          };
          content += `${contentItem.answer}[${position}]<br>`;
          position++;
          citations.push(clientCitation);
        } else {
          content += `${contentItem.answer}<br>`;
        }
      });
    } else {
      content = data.response.content;
    }

    if (isTeamsChannel) {
      content = content.split("\n").join("<br>");
    }

    // If there are citations, modify the content so that the sources are numbers instead of [doc1], [doc2], etc.
    const contentText = citations.length < 1 ? content : Utilities.formatCitationsResponse(content);

    // If there are citations, filter out the citations unused in content.
    const referencedCitations =
      citations.length > 0 ? Utilities.getUsedCitations(contentText, citations) : undefined;

    await context.sendActivity({
      type: ActivityTypes.Message,
      text: contentText,
      ...(isTeamsChannel ? { channelData: { feedbackLoopEnabled } } : {}),
      entities: [
        {
          type: "https://schema.org/Message",
          "@type": "Message",
          "@context": "https://schema.org",
          "@id": "",
          additionalType: ["AIGeneratedContent"],
          ...(referencedCitations ? { citation: referencedCitations } : {}),
        },
      ] as AIEntity[],
    });

    return "";
  };
}
