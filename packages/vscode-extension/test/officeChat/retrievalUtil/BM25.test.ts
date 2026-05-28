import * as chai from "chai";
import { BM25, DocumentWithmetadata } from "../../../src/officeChat/retrievalUtil/BM25";

const expect = chai.expect;

describe("BM25", function () {
  let bm25: BM25;
  let bm25WithConfig: BM25;
  let documents: DocumentWithmetadata[];
  let documents2: DocumentWithmetadata[];

  beforeEach(function () {
    documents = [
      { documentText: "This is a test document", metadata: null },
      { documentText: "Another test document", metadata: null },
      { documentText: "Yet another test document", metadata: null },
    ];

    documents2 = [
      { documentText: "", metadata: null },
      { documentText: "Another test document", metadata: null },
      { documentText: "Yet another test document", metadata: null },
    ];

    bm25 = new BM25(documents);
    bm25WithConfig = new BM25(documents2, { b: 0.5, k1: 1.5, d: 0.5, k3: 0.5 });
  });

  it("should create an instance", function () {
    expect(bm25).to.be.instanceOf(BM25);
    expect(bm25WithConfig).to.be.instanceOf(BM25);
  });

  it("should calculate average document length", function () {
    expect(bm25.averageLength).to.equal(4);
  });

  it("should perform a search", function () {
    const results = bm25.search(["test"]);
    expect(results).to.have.lengthOf(3);
    expect(results[0].score).to.be.above(0);

    const results2 = bm25WithConfig.search(["yet"]);
    expect(results2).to.have.lengthOf(3);
    expect(results2[0].score).to.be.above(0);
  });
});
