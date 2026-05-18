// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import "mocha";
import { parseAuthor } from "../../../../src/component/generator/openPlugin/authorParser";

describe("openPlugin.parseAuthor", () => {
  it("returns empty for null / undefined", () => {
    expect(parseAuthor(undefined)).to.deep.equal({});
    expect(parseAuthor(null)).to.deep.equal({});
  });

  it("parses 'Name <email> (url)' form", () => {
    expect(parseAuthor("Jane Doe <jane@example.com> (https://example.com)")).to.deep.equal({
      name: "Jane Doe",
      email: "jane@example.com",
      url: "https://example.com",
    });
  });

  it("parses 'Name <email>' form", () => {
    expect(parseAuthor("Jane Doe <jane@example.com>")).to.deep.equal({
      name: "Jane Doe",
      email: "jane@example.com",
    });
  });

  it("parses 'Name (url)' form", () => {
    expect(parseAuthor("Jane Doe (https://example.com)")).to.deep.equal({
      name: "Jane Doe",
      url: "https://example.com",
    });
  });

  it("parses bare name", () => {
    expect(parseAuthor("Jane Doe")).to.deep.equal({ name: "Jane Doe" });
  });

  it("parses object form", () => {
    expect(
      parseAuthor({ name: "Jane", email: "jane@example.com", url: "https://example.com" })
    ).to.deep.equal({
      name: "Jane",
      email: "jane@example.com",
      url: "https://example.com",
    });
  });

  it("ignores non-string fields on object form", () => {
    expect(parseAuthor({ name: 42 as unknown as string })).to.deep.equal({});
  });
});
