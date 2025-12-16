import { assert } from "chai";
import "mocha";
import fs from "fs-extra";
import sinon from "sinon";
import { AppManifestUtils } from "../src";
import * as nodeFetch from "node-fetch";

describe("AppManifestUtils", async () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("fetchSchema", async () => {
    it("should return local schema", async () => {
      const readJson = sandbox.stub(fs, "readJson").resolves({});
      const fetchStub = sandbox.stub(nodeFetch, "default").resolves({
        ok: true,
        json: async () => ({}),
      } as any);
      const schema = await AppManifestUtils.fetchSchema(
        "https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json"
      );
      assert.isTrue(readJson.calledOnce);
      assert.isTrue(fetchStub.notCalled);
      assert.deepEqual(schema, {} as any);
    });
    it("should fetch remote schema", async () => {
      const readJson = sandbox.stub(fs, "readJson").resolves({});
      const pathExists = sandbox.stub(fs, "pathExists").resolves(false);
      const fetchStub = sandbox.stub(nodeFetch, "default").resolves({
        ok: true,
        text: async () => JSON.stringify({}),
      } as any);
      const schema = await AppManifestUtils.fetchSchema("https://abc.schema.json");
      assert.isTrue(readJson.notCalled);
      assert.isTrue(fetchStub.calledOnce);
      assert.deepEqual(schema, {} as any);
    });
    it("should apply regex workaround for \\a and \\v characters when fetching remote schema", async () => {
      const mockResponseText = '{"pattern":"\\\\a test \\\\v pattern"}';
      const mockResponse = {
        text: sandbox.stub().resolves(mockResponseText),
      };
      const fetchStub = sandbox.stub(nodeFetch, "default").resolves(mockResponse as any);
      const readJson = sandbox.stub(fs, "readJson").rejects(new Error("File not found"));
      const pathExists = sandbox.stub(fs, "pathExists").resolves(false);
      const schema = await AppManifestUtils.fetchSchema(
        "https://developer.microsoft.com/json-schemas/teams/v1.24/MicrosoftTeams.schema.json"
      );
      assert.isTrue(fetchStub.calledOnce);
      assert.isTrue(mockResponse.text.calledOnce);
    });
  });
});
