import * as chai from "chai";
import sinon from "ts-sinon";
import { Tokenizer } from "../../src/chat/tokenizer";

describe("Tokenizer", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("getInstance", () => {
    const instance = Tokenizer.getInstance();
    chai.assert.isDefined(instance);
  });

  it("tokenize", () => {
    const tokenizer = new Tokenizer();
    const initStub = sandbox.stub(tokenizer as any, "initTokenize").returns({
      encode: () => [1, 2, 3],
    } as any);

    const result = tokenizer.tokenize("Hello world!");
    chai.assert.deepStrictEqual(result, [1, 2, 3]);
    chai.assert.isTrue(initStub.calledOnce);
  });

  describe("tokenLength", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("empty content", () => {
      const tokenizer = new Tokenizer();
      const result = tokenizer.tokenLength("");
      chai.assert.equal(result, 0);
    });

    it("non-empty content", () => {
      const tokenizer = new Tokenizer();
      const initStub = sandbox.stub(tokenizer as any, "initTokenize").returns({
        encode: () => [4, 5, 6],
      } as any);

      const result = tokenizer.tokenLength("Hello world!");
      chai.assert.equal(result, 3);
      chai.assert.isTrue(initStub.calledOnce);
    });
  });
});
