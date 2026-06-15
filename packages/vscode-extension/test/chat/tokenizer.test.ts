import * as chai from "chai";
import { Tokenizer } from "../../src/chat/tokenizer";
import { vi } from "vitest";

describe("Tokenizer", () => {
  it("getInstance", () => {
    const instance = Tokenizer.getInstance();
    chai.assert.isDefined(instance);
  });

  it("tokenize", () => {
    const tokenizer = new Tokenizer();
    const initStub = vi.spyOn(tokenizer as any, "initTokenize").mockReturnValue({
      encode: () => [1, 2, 3],
    } as any);

    const result = tokenizer.tokenize("Hello world!");
    chai.assert.deepStrictEqual(result, [1, 2, 3]);
    chai.assert.isTrue(initStub.calledOnce);
  });

  describe("tokenLength", () => {
    it("empty content", () => {
      const tokenizer = new Tokenizer();
      const result = tokenizer.tokenLength("");
      chai.assert.equal(result, 0);
    });

    it("non-empty content", () => {
      const tokenizer = new Tokenizer();
      const initStub = vi.spyOn(tokenizer as any, "initTokenize").mockReturnValue({
        encode: () => [4, 5, 6],
      } as any);

      const result = tokenizer.tokenLength("Hello world!");
      chai.assert.equal(result, 3);
      chai.assert.isTrue(initStub.calledOnce);
    });
  });
});
