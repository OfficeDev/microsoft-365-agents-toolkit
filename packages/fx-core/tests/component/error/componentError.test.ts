// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */
import { expect } from "chai";
import { BaseComponentInnerError } from "../../../src/component/error/componentError";
import { SystemError, UserError } from "@microsoft/teamsfx-api";

describe("BaseComponentInnerError", () => {
  describe("constructor", () => {
    it("should create a new BaseComponentInnerError with the correct properties", () => {
      const source = "test";
      const errorType = "UserError";
      const name = "TestError";
      const messageKey = "test.message";
      const messageParams = ["param1", "param2"];
      const suggestionKey = ["test.suggestion"];
      const detail = "Test error detail.";
      const helpLink = "https://example.com/help";
      const innerError = new Error("Inner error message.");
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        messageKey,
        messageParams,
        suggestionKey,
        detail,
        helpLink,
        innerError
      );
      expect(error).to.be.instanceOf(Error);
      expect(error).to.be.instanceOf(BaseComponentInnerError);
      expect(error.source).to.equal(source);
      expect(error.errorType).to.equal(errorType);
      expect(error.name).to.equal(name);
      expect(error.innerError).to.equal(innerError);
    });

    it("should create a new BaseComponentInnerError with the correct properties when messageKey is not provided", () => {
      const source = "test";
      const errorType = "UserError";
      const name = "TestError";
      const suggestionKey = ["test.suggestion"];
      const detail = "Test error detail.";
      const helpLink = "https://example.com/help";
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        undefined,
        undefined,
        suggestionKey,
        detail,
        helpLink,
        undefined
      );
      expect(error).to.be.instanceOf(Error);
      expect(error).to.be.instanceOf(BaseComponentInnerError);
      expect(error.source).to.equal(source);
      expect(error.errorType).to.equal(errorType);
      expect(error.name).to.equal(name);
      expect(error.displayMessage).to.equal("");
      expect(error.suggestionKey).to.deep.equal(suggestionKey);
      expect(error.detail).to.equal(detail);
      expect(error.helpLink).to.equal(helpLink);
      expect(error.innerError).to.equal(undefined);
    });

    it("should create a new BaseComponentInnerError with the correct properties when suggestionKey is not provided", () => {
      const source = "test";
      const errorType = "UserError";
      const name = "TestError";
      const messageKey = "test.message";
      const messageParams = ["param1", "param2"];
      const detail = "Test error detail.";
      const helpLink = "https://example.com/help";
      const innerError = new Error("Inner error message.");
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        messageKey,
        messageParams,
        undefined,
        detail,
        helpLink,
        innerError
      );
      expect(error).to.be.instanceOf(Error);
      expect(error).to.be.instanceOf(BaseComponentInnerError);
      expect(error.source).to.equal(source);
      expect(error.errorType).to.equal(errorType);
      expect(error.name).to.equal(name);
      expect(error.suggestionKey).to.be.undefined;
      expect(error.detail).to.equal(detail);
      expect(error.helpLink).to.equal(helpLink);
      expect(error.innerError).to.equal(innerError);
    });

    it("should create a new BaseComponentInnerError with the correct properties when detail is not provided", () => {
      const source = "test";
      const errorType = "UserError";
      const name = "TestError";
      const messageKey = "test.message";
      const messageParams = ["param1", "param2"];
      const suggestionKey = ["test.suggestion"];
      const helpLink = "https://example.com/help";
      const innerError = new Error("Inner error message.");
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        messageKey,
        messageParams,
        suggestionKey,
        undefined,
        helpLink,
        innerError
      );
      expect(error).to.be.instanceOf(Error);
      expect(error).to.be.instanceOf(BaseComponentInnerError);
      expect(error.source).to.equal(source);
      expect(error.errorType).to.equal(errorType);
      expect(error.name).to.equal(name);
      expect(error.message).to.equal("");
      expect(error.suggestionKey).to.deep.equal(suggestionKey);
      expect(error.detail).to.be.undefined;
      expect(error.helpLink).to.equal(helpLink);
      expect(error.innerError).to.equal(innerError);
    });

    it("should create a new BaseComponentInnerError with the correct properties when innerError is not provided", () => {
      const source = "test";
      const errorType = "UserError";
      const name = "TestError";
      const messageKey = "test.message";
      const messageParams = ["param1", "param2"];
      const suggestionKey = ["test.suggestion"];
      const detail = "Test error detail.";
      const helpLink = "https://example.com/help";
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        messageKey,
        messageParams,
        suggestionKey,
        detail,
        helpLink
      );
      expect(error).to.be.instanceOf(Error);
      expect(error).to.be.instanceOf(BaseComponentInnerError);
      expect(error.source).to.equal(source);
      expect(error.errorType).to.equal(errorType);
      expect(error.name).to.equal(name);
      expect(error.suggestionKey).to.deep.equal(suggestionKey);
      expect(error.detail).to.equal(detail);
      expect(error.helpLink).to.equal(helpLink);
      expect(error.innerError).to.be.undefined;
    });
  });

  describe("toFxError", () => {
    it("should return a new UserError with the correct properties when errorType is UserError", () => {
      const source = "test";
      const errorType = "UserError";
      const name = "TestError";
      const messageKey = "test.message";
      const messageParams = ["param1", "param2"];
      const suggestionKey = ["test.suggestion"];
      const detail = "Test error detail.";
      const helpLink = "https://example.com/help";
      const innerError = new Error("Inner error message.");
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        messageKey,
        messageParams,
        suggestionKey,
        detail,
        helpLink,
        innerError
      );
      const fxError = error.toFxError();
      expect(fxError).to.be.instanceOf(Error);
      expect(fxError).to.be.instanceOf(UserError);
      expect(fxError.source).to.equal(source);
      expect(fxError.name).to.equal(name);
      expect(fxError.message).to.equal("Inner error message.");
    });

    it("should return a new SystemError with the correct properties when errorType is SystemError", () => {
      const source = "test";
      const errorType = "SystemError";
      const name = "TestError";
      const messageKey = "test.message";
      const messageParams = ["param1", "param2"];
      const suggestionKey = ["test.suggestion"];
      const detail = "Test error detail.";
      const helpLink = "https://example.com/help";
      const error = new BaseComponentInnerError(
        source,
        errorType,
        name,
        messageKey,
        messageParams,
        suggestionKey,
        detail,
        helpLink,
        undefined
      );
      const fxError = error.toFxError();
      expect(fxError).to.be.instanceOf(Error);
      expect(fxError).to.be.instanceOf(SystemError);
      expect(fxError.source).to.equal(source);
      expect(fxError.name).to.equal(name);
      expect(fxError.innerError).to.equal(error);
    });
  });

  it("unknown error type should throw", () => {
    const error = BaseComponentInnerError.unknownError("test", "unknownErrorType");
    expect(error).to.be.instanceOf(BaseComponentInnerError);
    expect(error.source).to.equal("test");
    expect(error.errorType).to.equal("SystemError");
    expect(error.name).to.equal("UnhandledError");
    expect(error.innerError).to.be.undefined;
  });
});
