import { assert } from "chai";
import {
  CloudAdapter,
  ConversationReference,
  StatusCodes,
  TurnContext,
} from "@microsoft/agents-hosting";
import * as sinon from "sinon";
import { CardActionBot } from "../../../../src/conversationWithCloudAdapter/cardAction";
import { CardActionMiddleware } from "../../../../src/conversation/middlewares/cardActionMiddleware";
import {
  MockActionInvokeContext,
  MockCardActionHandler,
  MockCardActionHandlerWithErrorResponse,
} from "../conversation/testUtils";
import { InvokeResponseErrorCode } from "../../../../src/conversation/interface";

describe("Card Action Handler - Node", () => {
  it("handler should send text message response correctly", async () => {
    const doStuffAction = new MockCardActionHandler("doStuff", "sample-response");
    const testContext = new MockActionInvokeContext("doStuff");
    const middleware = new CardActionMiddleware([doStuffAction]);
    await middleware.onTurn(testContext as any, async () => {});

    assert.isTrue(doStuffAction.isInvoked);
    assert.strictEqual(testContext.content, "sample-response");
  });

  it("handler should send adaptive card response correctly", async () => {
    const responseCard = {
      version: "1.0.0",
      type: "AdaptiveCard",
      body: [
        {
          type: "TextBlock",
          text: `Hello World!`,
        },
      ],
    };

    const doStuffAction = new MockCardActionHandler("doStuff", responseCard);
    const testContext = new MockActionInvokeContext("doStuff");
    const middleware = new CardActionMiddleware([doStuffAction]);
    await middleware.onTurn(testContext as any, async () => {});

    assert.isTrue(doStuffAction.isInvoked);
    assert.deepEqual(testContext.content, responseCard);
  });

  it("handler should send user error response correctly", async () => {
    const errorMessage = "Invalid request";
    const doStuffAction = new MockCardActionHandlerWithErrorResponse(
      "doStuff",
      InvokeResponseErrorCode.BadRequest,
      errorMessage
    );
    const testContext = new MockActionInvokeContext("doStuff");
    const middleware = new CardActionMiddleware([doStuffAction]);
    await middleware.onTurn(testContext as any, async () => {});

    assert.isTrue(doStuffAction.isInvoked);
    assert.isNotNull(testContext.content);
    assert.strictEqual(testContext.content.message, errorMessage);
    assert.strictEqual(testContext.content.code, StatusCodes.BAD_REQUEST.toString());
  });

  it("handler should send server error response correctly", async () => {
    const errorMessage = "Internal server error";
    const doStuffAction = new MockCardActionHandlerWithErrorResponse(
      "doStuff",
      InvokeResponseErrorCode.InternalServerError,
      errorMessage
    );
    const testContext = new MockActionInvokeContext("doStuff");
    const middleware = new CardActionMiddleware([doStuffAction]);
    await middleware.onTurn(testContext as any, async () => {});

    assert.isTrue(doStuffAction.isInvoked);
    assert.isNotNull(testContext.content);
    assert.strictEqual(testContext.content.message, errorMessage);
    assert.strictEqual(testContext.content.code, StatusCodes.INTERNAL_SERVER_ERROR.toString());
  });

  it("handler should get action data correctly", async () => {
    const doStuffAction = new MockCardActionHandler("doStuff", "sampleResponse");
    const testContext = new MockActionInvokeContext("doStuff", { foo: "bar" });
    const middleware = new CardActionMiddleware([doStuffAction]);
    await middleware.onTurn(testContext as any, async () => {});

    assert.isTrue(doStuffAction.isInvoked);
    assert.deepEqual(doStuffAction.actionData, { foo: "bar" });
  });
});

describe("ard Action Bot Tests - Node", () => {
  const sandbox = sinon.createSandbox();
  let adapter: CloudAdapter;
  let middlewares: any[];

  beforeEach(() => {
    middlewares = [];
    const stubContext = sandbox.createStubInstance(TurnContext);
    const stubAdapter = sandbox.createStubInstance(CloudAdapter);
    stubAdapter.use.callsFake((args) => {
      middlewares.push(args);
      return stubAdapter;
    });
    (
      stubAdapter.continueConversation as unknown as sinon.SinonStub<
        [Partial<ConversationReference>, (context: TurnContext) => Promise<void>],
        Promise<void>
      >
    ).callsFake(async (ref, logic) => {
      await logic(stubContext);
    });
    adapter = stubAdapter;
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("initialize cardAction should create correct middleware", () => {
    const cardAction = new CardActionBot(adapter);
    assert.strictEqual(middlewares.length, 1);
    assert.isTrue(middlewares[0] instanceof CardActionMiddleware);
  });

  it("registerHandler should add card action successfully", () => {
    const cardAction = new CardActionBot(adapter);
    cardAction.registerHandler(new MockCardActionHandler("myAction"));

    assert.strictEqual(middlewares.length, 1);
    assert.isTrue(middlewares[0] instanceof CardActionMiddleware);
    const middleware = middlewares[0] as CardActionMiddleware;

    assert.isNotEmpty(middleware.actionHandlers);
    assert.isTrue(middleware.actionHandlers.length === 1);
    assert.isTrue(middleware.actionHandlers[0] instanceof MockCardActionHandler);
  });

  it("registerHandlers should add card actions successfully", () => {
    const cardAction = new CardActionBot(adapter);
    cardAction.registerHandlers([
      new MockCardActionHandler("myAction1"),
      new MockCardActionHandler("myAction2"),
    ]);

    assert.strictEqual(middlewares.length, 1);
    assert.isTrue(middlewares[0] instanceof CardActionMiddleware);
    const middleware = middlewares[0] as CardActionMiddleware;

    assert.isNotEmpty(middleware.actionHandlers);
    assert.isTrue(middleware.actionHandlers.length === 2);
    assert.isTrue(middleware.actionHandlers[0] instanceof MockCardActionHandler);
    assert.isTrue(middleware.actionHandlers[0].triggerVerb == "myAction1");
    assert.isTrue(middleware.actionHandlers[1] instanceof MockCardActionHandler);
    assert.isTrue(middleware.actionHandlers[1].triggerVerb == "myAction2");
  });
});
