const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const helloWorldCard = require("./adaptiveCards/helloWorldCard.json");

class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
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

module.exports.SearchApp = SearchApp;
