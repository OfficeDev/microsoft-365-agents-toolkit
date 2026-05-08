// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import * as fs from "fs-extra";
import * as path from "path";
import * as sinon from "sinon";
import proxyquire from "proxyquire";
chai.use(chaiAsPromised);
describe("portChecker", () => {
  const projectPath = path.resolve(__dirname, "data");

  describe("getPortsInUse()", () => {
    beforeEach(() => {
      sinon.restore();
      fs.ensureDirSync(projectPath);
      fs.emptyDirSync(projectPath);
    });

    afterEach(() => {
      sinon.restore();
    });

    it("happy path", async () => {
      const portChecker = proxyquire("../../../src/component/local/portChecker", {
        "detect-port": async (port: number) => port,
      });

      const waitingCheckPorts = [53000, 3978];
      const ports = await portChecker.getPortsInUse(waitingCheckPorts);

      chai.assert.isDefined(ports);
      chai.assert.equal(ports.length, 0);
    });

    it("detect-port timeout", async () => {
      const portChecker = proxyquire("../../../src/component/local/portChecker", {
        "detect-port": async (port: number) =>
          new Promise((resolve) => {
            setTimeout(() => resolve(port + 1), 60 * 1000);
          }),
      });
      const clock = sinon.useFakeTimers();

      try {
        const waitingCheckPorts = [3978];
        const portsPromise = portChecker.getPortsInUse(waitingCheckPorts);
        clock.tick(30 * 1000);
        const ports = await portsPromise;

        chai.assert.isDefined(ports);
        chai.assert.equal(ports.length, 0);
      } finally {
        clock.restore();
      }
    });

    it("53000 in use", async () => {
      const portChecker = proxyquire("../../../src/component/local/portChecker", {
        "detect-port": async (port: number) => (port === 53000 ? 53001 : port),
      });

      const waitingCheckPorts = [53000, 3978];
      const ports = await portChecker.getPortsInUse(waitingCheckPorts);

      chai.assert.isDefined(ports);
      chai.assert.deepEqual(ports, [53000]);
    });

    it("55000 in use, do not detect", async () => {
      const portChecker = proxyquire("../../../src/component/local/portChecker", {
        "detect-port": async (port: number) => (port === 55000 ? 55001 : port),
      });

      const waitingCheckPorts = [53000, 3978];
      const ports = await portChecker.getPortsInUse(waitingCheckPorts);

      chai.assert.isDefined(ports);
      chai.assert.deepEqual(ports, []);
    });

    it("dev:teamsfx port", async () => {
      const content = `\
        {\n\
          "name": "test",\n\
          "version": "1.0.0",\n\
          "scripts": {\n\
            "dev:teamsfx": "npm run dev",\n\
            "dev": "npx func start --inspect = '9229'"\n\
          }\n\
        }`;
      const packageJsonPath = path.join(projectPath, "api/package.json");
      await fs.ensureDir(path.join(projectPath, "api"));
      await fs.writeFile(packageJsonPath, content);

      const portChecker = proxyquire("../../../src/component/local/portChecker", {
        "detect-port": async (port: number) => (port === 9229 ? 9230 : port),
      });

      const waitingCheckPorts = [3978, 9229, 9239];
      const ports = await portChecker.getPortsInUse(waitingCheckPorts);

      chai.assert.isDefined(ports);
      chai.assert.deepEqual(ports, [9229]);
    });
  });
});
