import * as chai from "chai";
import * as fs from "fs-extra";
import mockFs from "mock-fs";
import * as sinon from "sinon";
import proxyquire from "proxyquire";
import * as projectSettingsHelper from "../../src/common/projectSettingsHelper";
import { OfficeManifestType } from "../../src/common/projectSettingsHelper";

describe("validateIsOfficeAddInProject", () => {
  const sandbox = sinon.createSandbox();
  let fetchManifestListStub: any;

  beforeEach(() => {
    fetchManifestListStub = sinon.stub(projectSettingsHelper, "fetchManifestList");
  });

  afterEach(() => {
    fetchManifestListStub.restore();
    mockFs.restore();
    sandbox.restore();
  });

  it("should return true if manifest list is not empty", () => {
    fetchManifestListStub.callsFake((workspace: string, type: OfficeManifestType) => {
      if (type == OfficeManifestType.XmlAddIn) {
        return ["manifest.xml"];
      } else {
        return [];
      }
    });
    mockFs({
      "/test/manifest.xml": "",
    });
    chai.expect(projectSettingsHelper.isValidOfficeAddInProject("/test")).to.be.true;
  });

  it("should return false if no manifest file", () => {
    fetchManifestListStub.returns([]);
    mockFs({
      "/test/useless.xml": "",
    });
    chai.expect(projectSettingsHelper.isValidOfficeAddInProject("/test")).to.be.false;
  });

  it("should return false if fetchManifestList throws an error", () => {
    fetchManifestListStub.throws(new Error("Error fetching manifest list"));
    chai.expect(projectSettingsHelper.isValidOfficeAddInProject("")).to.be.false;
  });

  it("should return false if both manifest.xml and manifest.json exist", () => {
    fetchManifestListStub.callsFake((workspace: string, type: OfficeManifestType) => {
      if (type == OfficeManifestType.XmlAddIn) {
        return ["manifest.xml"];
      } else if (type == OfficeManifestType.MetaOsAddIn) {
        return ["manifest.json"];
      } else {
        return [];
      }
    });
    mockFs({
      "/test/manifest.xml": "",
      "/test/manifest.json": "",
    });
    chai.expect(projectSettingsHelper.isValidOfficeAddInProject("/test")).to.be.false;
  });
});

describe("fetchManifestList", () => {
  let readdirSyncStub: sinon.SinonStub;
  let proxiedHelper: typeof projectSettingsHelper;

  beforeEach(() => {
    readdirSyncStub = sinon.stub();
    proxiedHelper = proxyquire("../../src/common/projectSettingsHelper", {
      "fs-extra": {
        ...fs,
        readdirSync: readdirSyncStub,
      },
    });
  });

  afterEach(() => {
    sinon.restore();
    mockFs.restore();
  });

  it("should return undefined if workspacePath is not provided", () => {
    chai.expect(proxiedHelper.fetchManifestList()).to.be.undefined;
  });

  it("should return manifest.xml if type is OfficeManifestType.XmlAddIn", () => {
    mockFs({
      "/test/manifest.xml": "",
    });
    readdirSyncStub.returns(["manifest.xml"]);
    chai
      .expect(proxiedHelper.fetchManifestList("/test", proxiedHelper.OfficeManifestType.XmlAddIn))
      .to.deep.equal(["manifest.xml"]);
  });

  it("should return manifest.json if type is OfficeManifestType.MetaOsAddIn", () => {
    mockFs({
      "/test/manifest.json": "",
    });
    readdirSyncStub.returns(["manifest.json"]);
    chai
      .expect(
        proxiedHelper.fetchManifestList("/test", proxiedHelper.OfficeManifestType.MetaOsAddIn)
      )
      .to.deep.equal(["manifest.json"]);
  });

  it("should return false if both manifest.xml and manifest.json exist but type is OfficeManifestType.XmlAddIn", () => {
    mockFs({
      "/test/manifest.xml": "",
      "/test/manifest.json": "",
    });
    readdirSyncStub.returns(["manifest.xml", "manifest.json"]);
    chai
      .expect(proxiedHelper.fetchManifestList("/test", proxiedHelper.OfficeManifestType.XmlAddIn))
      .to.deep.equal(["manifest.xml"]);
  });

  it("should return true if manifest.json exist and type is OfficeManifestType.MetaOsAddIn", () => {
    mockFs({
      "/test/manifest.xml": "",
      "/test/manifest.json": "",
    });
    readdirSyncStub.returns(["manifest.xml", "manifest.json"]);
    chai
      .expect(
        proxiedHelper.fetchManifestList("/test", proxiedHelper.OfficeManifestType.MetaOsAddIn)
      )
      .to.deep.equal(["manifest.json"]);
  });

  it("should return true when no src folder exists", () => {
    mockFs({
      "/test/manifest.xml": "",
    });
    chai.expect(projectSettingsHelper.isManifestOnlyOfficeAddinProject("/test")).to.be.true;
  });
});
