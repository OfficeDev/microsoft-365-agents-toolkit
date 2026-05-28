import { expect } from "chai";
import {
  stemText,
  converseSynonym,
  prepareExamples,
} from "../../../src/officeChat/retrievalUtil/retrievalUtil";

describe("stemText", function () {
  it("should stem the input texts", function () {
    const input = ["running", "jumps", "happily"];
    const output = stemText(input);
    expect(output).to.deep.equal(["run", "jump", "happili"]);
  });

  it("should return the same word in synonymReplaceRules", function () {
    const input = "fetch";
    const output = converseSynonym(input);
    expect(output).to.deep.equal("get");
  });
});

describe("prepareExamples", function () {
  it("should prepare examples and return an array and a map", function () {
    const docs = [
      { description: "This is a test document", codeSample: 'console.log("Hello, world!")' },
      { description: "Another test document", codeSample: 'console.log("Hello, again!")' },
    ];
    const [cleanDocs, docsWithMetadata] = prepareExamples(docs);
    expect(cleanDocs).to.deep.equal(["This test docum", "Another test docum"]);
    expect(docsWithMetadata.get("This test docum")).to.deep.equal(docs[0]);
    expect(docsWithMetadata.get("Another test docum")).to.deep.equal(docs[1]);
  });
});
