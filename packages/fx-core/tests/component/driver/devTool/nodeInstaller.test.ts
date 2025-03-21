// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import chai from "chai";
import "mocha";
import * as nodeFetch from "node-fetch";
import { Response } from "node-fetch";
import * as sinon from "sinon";
import { httpClient } from "../../../../src/component/driver/devTool/httpClient";
import { ToolsInstallDriver } from "../../../../src/component/driver/devTool/installDriver";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { Readable } from "stream";

describe("NodeJS Installer", () => {
  const sandbox = sinon.createSandbox();
  const toolsInstallDriver = new ToolsInstallDriver();
  const mockedDriverContext: any = {
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
  };

  describe("HttpClient", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("fetch return 500", async () => {
      sandbox.stub(nodeFetch, "default").resolves({ ok: false, status: 500 } as any);
      try {
        await httpClient.get("https://text.com");
      } catch (e: any) {
        chai.assert.equal(e.message, "Request failed with status 500");
      }
    });

    it("happy", async () => {
      const bufferChunks = [Buffer.from("chunk1"), Buffer.from("chunk2")];
      const fakeResponse = new Response(Readable.from(bufferChunks), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
      sandbox.stub(nodeFetch, "default").resolves(fakeResponse);
      const result = await httpClient.get("https://text.com");
      chai.assert.equal(result, "chunk1chunk2");
    });
  });
});
