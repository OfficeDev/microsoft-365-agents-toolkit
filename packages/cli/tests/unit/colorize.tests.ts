// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import sinon from "sinon";
import { TextType, colorize, replaceTemplateString } from "../../src/colorize";
import ScreenManager from "../../src/console/screen";

describe("colorize", () => {
  const sandox = sinon.createSandbox();
  let message = "";
  let isTTY: boolean;
  beforeEach(() => {
    isTTY = process.stdout.isTTY;
    process.stdout.isTTY = true;
    sandox.stub(ScreenManager, "writeLine").callsFake((msg: string) => (message += msg));
  });

  afterEach(() => {
    process.stdout.isTTY = isTTY;
    sandox.restore();
    message = "";
  });

  it("colorize - Success", async () => {
    colorize("test", TextType.Success);
  });

  it("colorize - Error", async () => {
    colorize("test", TextType.Error);
  });

  it("colorize - Warning", async () => {
    colorize("test", TextType.Warning);
  });

  it("colorize - Info", async () => {
    colorize("test", TextType.Info);
  });

  it("colorize - Hyperlink", async () => {
    colorize("test", TextType.Hyperlink);
  });

  it("colorize - Email", async () => {
    colorize("test", TextType.Email);
  });

  it("colorize - Important", async () => {
    colorize("test", TextType.Important);
  });

  it("colorize - Details", async () => {
    colorize("test", TextType.Details);
  });
  it("colorize - Commands", async () => {
    colorize("test", TextType.Commands);
  });
  it("replace template string", async () => {
    const template = "test %s";
    const result = replaceTemplateString(template, "test");
    expect(result).to.equal("test test");
  });
});
