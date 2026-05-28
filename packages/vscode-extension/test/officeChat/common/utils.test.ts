import * as chai from "chai";
import * as sinon from "sinon";
import fs from "fs";
import chaiPromised from "chai-as-promised";
import * as commonUtils from "../../../src/officeChat/common/utils";
import * as requestUtils from "@microsoft/teamsfx-core/build/common/requestUtils";
import { AxiosResponse } from "axios";

chai.use(chaiPromised);

describe("File: officeChat/common/utils", () => {
  const sandbox = sinon.createSandbox();

  describe("Method: fetchRawFileContent", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("return file response data", async () => {
      sandbox.stub(requestUtils, "sendRequestWithTimeout").resolves({
        data: "testData",
      } as AxiosResponse);
      const result = await commonUtils.fetchRawFileContent("test");
      chai.assert.equal(result, "testData");
    });

    it("return empty string", async () => {
      sandbox.stub(requestUtils, "sendRequestWithTimeout").resolves(undefined);
      const result = await commonUtils.fetchRawFileContent("test");
      chai.assert.equal(result, "");
    });

    it("throw error", async () => {
      sandbox.stub(requestUtils, "sendRequestWithTimeout").rejects();
      try {
        await commonUtils.fetchRawFileContent("test");
        chai.assert.fail("should not reach here");
      } catch (e) {
        chai.assert.equal((e as Error).message, "Cannot fetch test.");
      }
    });
  });

  describe("Method: compressCode", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("compress code", () => {
      const code = `
      const x = 1; /* This is another comment */
      const y = 2;  // Another comment
      const z = x +   y; // Comment with extra spaces
    `;
      const expected = `
 const x = 1; 
 const y = 2; 
 const z = x + y; 
 `;
      const result = commonUtils.compressCode(code);
      chai.expect(result).to.equal(expected);
    });
  });

  describe("Method: sleep helpers", () => {
    let clock: any;
    let randomStub: any;

    beforeEach(() => {
      clock = sandbox.useFakeTimers();
      randomStub = sandbox.stub(Math, "random");
    });

    afterEach(() => {
      clock.restore();
      randomStub.restore();
      sandbox.restore();
    });

    it("sleep", async () => {
      const promise = commonUtils.sleep(5);
      clock.tick(5000); // Advance the timer by 5 seconds
      await promise; // This should now resolve immediately
      chai.expect(clock.now).to.equal(5000); // Check that 5 seconds have passed
    });

    it("sleepRandom", async () => {
      const minSecond = 1;
      const maxSecond = 5;
      randomStub.returns(0); // This will make Math.random() always return 0, so the sleep time will always be minSecond
      const promise = commonUtils.sleepRandom(minSecond, maxSecond);
      clock.tick(minSecond * 1000); // Advance the timer by minSecond seconds
      await promise; // This should now resolve immediately
      chai.expect(clock.now).to.equal(minSecond * 1000); // Check that minSecond seconds have passed
    });
  });

  describe("Method: writeLogToFile", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("write log to file", async () => {
      const appendFileStub = sandbox.stub(fs, "appendFileSync");
      await commonUtils.writeLogToFile("test");
      chai.assert.isTrue(appendFileStub.calledOnceWith("C:\\temp\\codeGenLog.txt", "test"));
    });
  });

  describe("Method: correctPropertyLoadSpelling", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("should correct the spelling of property load", () => {
      const codeSnippet = `chart.load("name", "chartType", "height", "width");`;
      const expected = `chart.load("name, chartType, height, width");`;
      const result = commonUtils.correctPropertyLoadSpelling(codeSnippet);
      chai.expect(result).to.equal(expected);
    });
  });
});
