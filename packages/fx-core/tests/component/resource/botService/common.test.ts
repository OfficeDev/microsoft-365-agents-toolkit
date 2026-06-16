// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as chai from "chai";
import * as utils from "../../../../src/component/resource/botService/common";

describe("Common Utils", () => {
  describe("isHttpCodeOkOrCreated", () => {
    it("Expect True", () => {
      // Arrange
      const code = 200;

      // Act
      const result = utils.isHttpCodeOkOrCreated(code);

      // Assert
      chai.assert.isTrue(result);
    });
  });
});
