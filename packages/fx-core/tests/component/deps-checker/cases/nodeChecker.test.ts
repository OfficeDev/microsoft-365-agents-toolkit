import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import * as path from "path";
import * as sinon from "sinon";
import { DepsType, EmptyLogger, EmptyTelemetry } from "../../../../src/component/deps-checker";
import {
  NodeChecker,
  ProjectNodeChecker,
} from "../../../../src/component/deps-checker/internal/nodeChecker";

chai.use(chaiAsPromised);

const nodeVersionFolder = path.resolve(__dirname, "../data/node-version/");
const noNodeVersionFolder = path.resolve(__dirname, "../data/node-version/");
describe("NodeChecker", () => {
  describe("ProjectNodeChecker", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("No node engine property specified", async () => {
      const nodeChecker = new ProjectNodeChecker(new EmptyLogger(), new EmptyTelemetry());

      sandbox
        .stub(NodeChecker, "getInstalledNodeVersion")
        .resolves({ version: "v14.17.6", majorVersion: "14" });

      const res = await nodeChecker.resolve({ projectPath: noNodeVersionFolder });
      chai.assert.deepEqual(res, {
        command: "node",
        details: {
          installVersion: "v14.17.6",
          isLinuxSupported: true,
          supportedVersions: [">=14 <=16"],
        },
        error: undefined,
        isInstalled: true,
        name: "Node.js",
        type: DepsType.ProjectNode,
      });
    });

    it("Lower node version", async () => {
      const nodeChecker = new ProjectNodeChecker(new EmptyLogger(), new EmptyTelemetry());

      sandbox
        .stub(NodeChecker, "getInstalledNodeVersion")
        .resolves({ version: "v12.20.0", majorVersion: "12" });

      const res = await nodeChecker.resolve({ projectPath: nodeVersionFolder });
      chai.assert.equal(res.command, "node");
      chai.assert.isDefined(res.details);
      chai.assert.equal(res.details.installVersion, "v12.20.0");
      chai.assert.isTrue(res.details.isLinuxSupported);
      chai.assert.sameMembers(res.details.supportedVersions, [">=14 <=16"]);

      chai.assert.isDefined(res.error);
      chai.assert.isTrue(res.isInstalled);
      chai.assert.equal(res.name, "Node.js");
      chai.assert.equal(res.type, DepsType.ProjectNode);
    });

    it("Higher node version", async () => {
      const nodeChecker = new ProjectNodeChecker(new EmptyLogger(), new EmptyTelemetry());

      sandbox
        .stub(NodeChecker, "getInstalledNodeVersion")
        .resolves({ version: "v18.9.0", majorVersion: "18" });

      const res = await nodeChecker.resolve({ projectPath: nodeVersionFolder });
      chai.assert.equal(res.command, "node");
      chai.assert.isDefined(res.details);
      chai.assert.equal(res.details.installVersion, "v18.9.0");
      chai.assert.isTrue(res.details.isLinuxSupported);
      chai.assert.sameMembers(res.details.supportedVersions, [">=14 <=16"]);

      chai.assert.isDefined(res.error);
      chai.assert.isTrue(res.isInstalled);
      chai.assert.equal(res.name, "Node.js");
      chai.assert.equal(res.type, DepsType.ProjectNode);
    });

    it("Supported node version", async () => {
      const nodeChecker = new ProjectNodeChecker(new EmptyLogger(), new EmptyTelemetry());

      sandbox
        .stub(NodeChecker, "getInstalledNodeVersion")
        .resolves({ version: "v14.17.5", majorVersion: "14" });

      const res = await nodeChecker.resolve({ projectPath: nodeVersionFolder });
      chai.assert.deepEqual(res, {
        command: "node",
        details: {
          installVersion: "v14.17.5",
          isLinuxSupported: true,
          supportedVersions: [">=14 <=16"],
        },
        error: undefined,
        isInstalled: true,
        name: "Node.js",
        type: DepsType.ProjectNode,
      });
    });
  });
});
