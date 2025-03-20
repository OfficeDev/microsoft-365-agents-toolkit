// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { createSandbox } from "sinon";
import { RetryHandler, GraphAPIClient, SensitivityLabel } from "../../src/client/graphAPIClient";
import { ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import "mocha";

describe("GraphAPIClient Test", () => {
  const sandbox = createSandbox();
  const token = "fakeToken";

  beforeEach(() => {
    sandbox.stub(RetryHandler, "RETRIES").value(1);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("RetryHandler", () => {
    it("Happy path", async () => {
      const fn = sandbox.stub().resolves("success");
      const result = await RetryHandler.Retry(fn);
      expect(result).to.equal("success");
      expect(fn.calledOnce).to.be.true;
    });

    it("Retry on error and succeed", async () => {
      const fn = sandbox.stub();
      fn.onFirstCall().rejects(new Error("Failed"));
      fn.onSecondCall().resolves("success");

      // Set RETRIES to 2 for this test
      sandbox.stub(RetryHandler, "RETRIES").value(2);

      const result = await RetryHandler.Retry(fn);
      expect(result).to.equal("success");
      expect(fn.calledTwice).to.be.true;
    });

    it("Fail after all retries", async () => {
      const error = new Error("Failed");
      const fn = sandbox.stub().rejects(error);

      try {
        await RetryHandler.Retry(fn);
        expect.fail("Should have thrown error");
      } catch (e) {
        expect(e).to.equal(error);
      }

      expect(fn.calledOnce).to.be.true;
    });
  });

  describe("listSensitivityLabels", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          value: [
            {
              id: "label1",
              displayName: "General",
              name: "General Label",
              description: "General Label Description",
            },
            {
              id: "label2",
              displayName: "Confidential",
              name: "Confidential Label",
              description: "Confidential Label Description",
            },
          ],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphAPIClient();
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.length).to.equal(2);
        expect(result.value[0].id).to.equal("label1");
        expect(result.value[0].displayName).to.equal("General");
        expect(result.value[1].id).to.equal("label2");
        expect(result.value[1].displayName).to.equal("Confidential");
      }
    });

    it("Return error for empty response", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {};
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphAPIClient();
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("listSensitivityLabelsError");
      }
    });

    it("Return error for empty data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = { data: {} };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphAPIClient();
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("listSensitivityLabelsError");
      }
    });

    it("API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = new Error("API failed");
      sandbox.stub(fakeAxiosInstance, "get").rejects(error);
      sandbox.stub(RetryHandler, "Retry").rejects(error);

      const graphAPIClient = new GraphAPIClient();
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("listSensitivityLabelsError");
        expect(result.error.message).to.include("API failed");
      }
    });
  });

  describe("getGeneralSentivityLabelId", () => {
    it("Happy path", async () => {
      const graphAPIClient = new GraphAPIClient();

      const labels: SensitivityLabel[] = [
        {
          id: "general-id",
          displayName: "General",
          name: "General Label",
          description: "General Label Description",
        },
        {
          id: "confidential-id",
          displayName: "Confidential",
          name: "Confidential Label",
          description: "Confidential Label Description",
        },
      ];

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(labels));

      const result = await graphAPIClient.getGeneralSentivityLabelId(token);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.equal("general-id");
      }
    });

    it("No General label found", async () => {
      const graphAPIClient = new GraphAPIClient();

      const labels: SensitivityLabel[] = [
        {
          id: "confidential-id",
          displayName: "Confidential",
          name: "Confidential Label",
          description: "Confidential Label Description",
        },
      ];

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(labels));

      const result = await graphAPIClient.getGeneralSentivityLabelId(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("getGeneralSentivityLabelIdError");
      }
    });

    it("General label has no ID", async () => {
      const graphAPIClient = new GraphAPIClient();

      const labels: SensitivityLabel[] = [
        {
          displayName: "General",
          name: "General Label",
          description: "General Label Description",
        },
        {
          id: "confidential-id",
          displayName: "Confidential",
          name: "Confidential Label",
          description: "Confidential Label Description",
        },
      ];

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(labels));

      const result = await graphAPIClient.getGeneralSentivityLabelId(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("getGeneralSentivityLabelIdError");
      }
    });

    it("listSensitivityLabels returns error", async () => {
      const graphAPIClient = new GraphAPIClient();

      const fakeError = {
        name: "listSensitivityLabelsError",
        message: "API failed",
        source: "GraphAPI",
      };

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves({
        isErr: () => true,
        isOk: () => false,
        error: fakeError,
        value: undefined,
      } as any);

      const result = await graphAPIClient.getGeneralSentivityLabelId(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error).to.equal(fakeError);
      }
    });
  });
});
