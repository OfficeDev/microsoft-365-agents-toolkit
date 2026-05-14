/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as sinon from "sinon";
import { colorize, TextType } from "../../src/output/colorize";
import { printResult } from "../../src/output/formatter";

describe("CLI Output", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("colorize()", () => {
    it("should return raw message when not a TTY", () => {
      // In test env, stdout is typically not a TTY
      const result = colorize("hello", TextType.Success);
      // When not TTY, should return message as-is
      if (!process.stdout.isTTY) {
        expect(result).to.equal("hello");
      } else {
        // In TTY, chalk adds ANSI codes, but message is still embedded
        expect(result).to.include("hello");
      }
    });

    it("should handle all TextType values without throwing", () => {
      for (const type of Object.values(TextType)) {
        expect(() => colorize("test", type)).to.not.throw();
      }
    });
  });

  describe("printResult()", () => {
    it("should print JSON when format is json", () => {
      const logStub = sandbox.stub(console, "log");
      printResult({ key: "value" }, { format: "json" });
      expect(logStub.calledOnce).to.be.true;
      const output = logStub.getCall(0).args[0];
      expect(JSON.parse(output)).to.deep.equal({ key: "value" });
    });

    it("should print table for text format with array data", () => {
      const logStub = sandbox.stub(console, "log");
      printResult([{ name: "test", value: 42 }], { format: "text" });
      expect(logStub.called).to.be.true;
    });

    it("should print (no results) for empty array", () => {
      const logStub = sandbox.stub(console, "log");
      printResult([], { format: "text" });
      expect(logStub.calledWith("(no results)")).to.be.true;
    });
  });
});
