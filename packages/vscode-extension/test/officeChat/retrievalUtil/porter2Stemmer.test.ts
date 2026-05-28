import * as chai from "chai";
import { stemmer } from "../../../src/officeChat/retrievalUtil/porter2Stemmer";

const expect = chai.expect;

describe("stemmer", function () {
  it("should stem the input word", function () {
    const input = "running";
    const output = stemmer(input);
    expect(output).to.equal("run");
  });

  it("should return the same word if it cannot be stemmed", function () {
    const input = "test";
    const output = stemmer(input);
    expect(output).to.equal("test");
  });

  it("should stem all words", function () {
    const input = [
      "abaissiez",
      "abandoned",
      "sky",
      "nefarious",
      "regenerate",
      "a",
      "Yes",
      "'yes",
      "coarseness",
      "generated",
      "cries",
      "excesses",
      "needly",
      "misdeed",
      "behaneedly",
      "cied",
    ];
    const expectedOutput = [
      "abaissiez",
      "abandon",
      "sky",
      "nefari",
      "regener",
      "a",
      "yes",
      "yes",
      "coars",
      "generat",
      "cri",
      "excess",
      "need",
      "misdee",
      "behane",
      "cie",
    ];
    const output = input.map(stemmer);
    expect(output).to.deep.equal(expectedOutput);
  });
});
