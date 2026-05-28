import { expect } from "chai";
import {
  OfficeTemplateModelPorvider,
  WXPAppName,
} from "../../../../src/officeChat/common/samples/officeTemplateModelPorvider";

describe("OfficeTemplateModelPorvider", () => {
  let provider: OfficeTemplateModelPorvider;

  beforeEach(() => {
    provider = OfficeTemplateModelPorvider.getInstance();
  });

  it("should return BM25Model PowerPoint", async () => {
    let bm25ModelPowerPoint = await provider.getBM25Model("PowerPoint");
    if (bm25ModelPowerPoint === null) {
      bm25ModelPowerPoint = await provider.getBM25Model("PowerPoint");
    }
    expect(bm25ModelPowerPoint).to.exist;

    const bm25ModelPowerPointCached = await provider.getBM25Model("PowerPoint");
    expect(bm25ModelPowerPointCached).to.equal(bm25ModelPowerPoint);
  }).timeout(5000);

  it("invalid host", async () => {
    const bm25ModelFake = await provider.getBM25Model("Fake" as WXPAppName);
    expect(bm25ModelFake).to.not.exist;

    const bm25ModelEmptyHost = await provider.getBM25Model("" as WXPAppName);
    expect(bm25ModelEmptyHost).to.not.exist;
  }).timeout(5000);
});
