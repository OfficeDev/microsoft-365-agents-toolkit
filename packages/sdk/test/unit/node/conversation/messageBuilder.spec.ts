import { assert } from "chai";
import { MessageBuilder } from "../../../../src";

describe("attachAdaptiveCard", () => {
  it("adaptive card with data", () => {
    const card = {
      type: "AdaptiveCard",
      body: [
        {
          type: "TextBlock",
          text: "Hello, ${name}!",
        },
      ],
    };

    const data = {
      name: "test",
    };

    const expectedCard = {
      attachments: [
        {
          content: {
            type: "AdaptiveCard",
            body: [
              {
                type: "TextBlock",
                text: "Hello, test!",
              },
            ],
          },
          contentType: "application/vnd.microsoft.card.adaptive",
        },
      ],
    };

    const result = MessageBuilder.attachAdaptiveCard(card, data);
    assert.deepStrictEqual(result, expectedCard);
  });

  it("adaptive card with without data", () => {
    const card = {
      type: "AdaptiveCard",
      body: [
        {
          type: "TextBlock",
          text: "Hello, test!",
        },
      ],
    };

    const expectedCard = {
      attachments: [
        {
          content: {
            type: "AdaptiveCard",
            body: [
              {
                type: "TextBlock",
                text: "Hello, test!",
              },
            ],
          },
          contentType: "application/vnd.microsoft.card.adaptive",
        },
      ],
    };

    const result = MessageBuilder.attachAdaptiveCardWithoutData(card);
    assert.deepStrictEqual(result, expectedCard);
  });

  it("signin card", () => {
    const title = "Sign in";
    const url = "https://example.com/signin";

    const expectedCard = {
      attachments: [
        {
          content: {
            buttons: [
              {
                type: "signin",
                title: title,
                value: url,
                channelData: undefined,
              },
            ],
          },
          contentType: "application/vnd.microsoft.card.signin",
        },
      ],
    };

    const result = MessageBuilder.attachSigninCard(title, url);
    assert.deepStrictEqual(result, expectedCard);
  });

  it("o365 connector card", () => {
    const card = {};

    const expectedCard = {
      attachments: [
        {
          content: {},
          contentType: "application/vnd.microsoft.teams.card.o365connector",
        },
      ],
    };

    const result = MessageBuilder.attachO365ConnectorCard(card);
    assert.deepStrictEqual(result, expectedCard);
  });

  it("Receipt card", () => {
    const card = {
      title: "Receipt",
      items: [],
      facts: [],
      total: "100",
      tax: "10",
      buttons: [],
    };

    const expectedCard = {
      attachments: [
        {
          content: card,
          contentType: "application/vnd.microsoft.card.receipt",
        },
      ],
    };

    const result = MessageBuilder.AttachReceiptCard(card);
    assert.deepStrictEqual(result, expectedCard);
  });

  it("thumbnail card", () => {
    const title = "Thumbnail Card";

    const expectedCard = {
      attachments: [
        {
          content: {
            title: title,
          },
          contentType: "application/vnd.microsoft.card.thumbnail",
        },
      ],
    };

    const result = MessageBuilder.attachThumbnailCard(title);
    assert.deepStrictEqual(result, expectedCard);
  });

  it("hero card", () => {
    const title = "Hero Card";

    const expectedCard = {
      attachments: [
        {
          content: {
            title: title,
          },
          contentType: "application/vnd.microsoft.card.hero",
        },
      ],
    };

    const result = MessageBuilder.attachHeroCard(title);
    assert.deepStrictEqual(result, expectedCard);
  });
});
