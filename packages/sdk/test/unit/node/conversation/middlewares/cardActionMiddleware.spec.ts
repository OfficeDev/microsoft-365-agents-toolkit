// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert, use as chaiUse } from "chai";
import * as chaiPromises from "chai-as-promised";
import { MockActionInvokeContext, MockCardActionHandler } from "../testUtils";
import { CardActionMiddleware } from "../../../../../src/conversation/middlewares/cardActionMiddleware";

chaiUse(chaiPromises);

describe("CardAction Middleware Tests - Node", () => {
  it("onTurn should invoke card action handler if verb is matched", async () => {
    const doStuffAction = new MockCardActionHandler("doStuff", "myResponseMessage");
    const middleware = new CardActionMiddleware([doStuffAction]);

    const testContext = new MockActionInvokeContext("doStuff");
    await middleware.onTurn(testContext as any, async () => {});

    // Assert the card action handler is invoked
    assert.isTrue(doStuffAction.isInvoked);
  });

  it("onTurn should invoke card action handler if verb is matched - adaptivecard", async () => {
    const card = {
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
      refresh: true,
    };
    const doStuffAction = new MockCardActionHandler("doStuff", card);
    const middleware = new CardActionMiddleware([doStuffAction]);

    const testContext = new MockActionInvokeContext("doStuff");
    await middleware.onTurn(testContext as any, async () => {});

    // Assert the card action handler is invoked
    assert.isTrue(doStuffAction.isInvoked);
  });

  it("onTurn shouldn't invoke card action handler if verb is not matched", async () => {
    const doStuffAction = new MockCardActionHandler("doStuff", "myResponseMessage");
    const middleware = new CardActionMiddleware([doStuffAction]);

    const testContext = new MockActionInvokeContext("inconsistent-verb");
    await middleware.onTurn(testContext as any, async () => {});

    // Assert the card action handler is not invoked
    assert.isFalse(doStuffAction.isInvoked);
  });
});
