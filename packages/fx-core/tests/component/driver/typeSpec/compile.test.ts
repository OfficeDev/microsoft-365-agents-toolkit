// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { MockedM365Provider, MockLogProvider, MockTools } from "../../../core/utils";
import { TypeSpecCompileArgs } from "../../../../src/component/driver/typeSpec/interface/typeSpecCompileArgs";
import { MockedUserInteraction } from "../../../plugins/solution/util";
import fs from "fs-extra";
import {
  DeclarativeCopilotManifestSchema,
  err,
  ok,
  Platform,
  SystemError,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import { TypeSpecCompileDriver } from "../../../../src/component/driver/typeSpec/compile";

chai.use(chaiAsPromised);
const expect = chai.expect;
const tools = new MockTools();
const mockedDriverContext: any = {
  m365TokenProvider: new MockedM365Provider(),
  ui: new MockedUserInteraction(),
  projectPath: "test",
  platform: Platform.VSCode,
  logProvider: new MockLogProvider(),
};
mockedDriverContext.logProvider.outputChannel = {
  show: () => {
    return;
  },
};

describe("typeSpecCompilt", async () => {
  const sandbox = sinon.createSandbox();
  let envRestore: RestoreFn | undefined;
  const typeSpecCompileDriver = new TypeSpecCompileDriver();
  const manifest: TeamsAppManifest = {
    manifestVersion: "mockedManifestVersion",
    version: "mockedVersion",
    id: "mockedId",
    developer: {
      name: "mockedName",
      websiteUrl: "mockedWebsiteUrl",
      privacyUrl: "mockedPrivacyUrl",
      termsOfUseUrl: "mockedTermsOfUseUrl",
    },
    name: {
      short: "mockedShortName",
      full: "mockedFullName",
    },
    description: {
      short: "mockedShortDescription",
      full: "mockedFullDescription",
    },
    icons: {
      outline: "mockedOutlineIcon",
      color: "mockedColorIcon",
    },
    accentColor: "mockedAccentColor",
  };

  beforeEach(() => {
    envRestore = mockedEnv({
      TEAMSFX_TYPESPEC: "true",
    });
  });

  afterEach(() => {
    sandbox.restore();
    if (envRestore) {
      envRestore();
    }
  });

  it("happy path: with one action", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId",
          file: "mockedFile",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox.stub(mockedDriverContext.ui, "runCommand").resolves(ok("mockedCommandResult"));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox
      .stub(fs, "readJSON")
      .onFirstCall()
      .resolves(pluginManifest)
      .onSecondCall()
      .resolves(manifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
  });

  it("happy path: with multiple actions", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedAction1",
          file: "mockedAction1-apiplugin.json",
        },
        {
          id: "mockedActionId2",
          file: "mockedAction2-apiplugin.json",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    const runCommandStub = sandbox
      .stub(mockedDriverContext.ui, "runCommand")
      .resolves(ok("mockedCommandResult"));
    sandbox.stub(fs, "readdirSync").returns(["mockedAction1.yaml", "mockedAction2.yaml"] as any);
    sandbox
      .stub(fs, "readJSON")
      .onFirstCall()
      .resolves(pluginManifest)
      .onSecondCall()
      .resolves(manifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    expect(runCommandStub.calledThrice).to.be.true;
  });

  it("should throw error if missing input", async () => {
    let result = await typeSpecCompileDriver.execute(
      {
        path: "mockedPath",
        manifestPath: "",
      },
      mockedDriverContext
    );
    expect(result.result.isErr()).to.be.true;

    result = await typeSpecCompileDriver.execute(
      {
        path: "",
        manifestPath: "mockedManifestPath",
      },
      mockedDriverContext
    );
    expect(result.result.isErr()).to.be.true;
  });

  it("should throw error if failed to run tsp command", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId",
          file: "mockedFile",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox
      .stub(mockedDriverContext.ui, "runCommand")
      .returns(err(new SystemError("mockedSource", "mockedError", "mockedErrorMessage")));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox
      .stub(fs, "readJSON")
      .onFirstCall()
      .resolves(pluginManifest)
      .onSecondCall()
      .resolves(manifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
  });

  it("should throw error if failed to run kiota command", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId",
          file: "mockedFile",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox
      .stub(mockedDriverContext.ui, "runCommand")
      .onFirstCall()
      .resolves(ok("mockedCommandResult"))
      .onSecondCall()
      .returns(err(new SystemError("mockedSource", "mockedError", "mockedErrorMessage")));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox
      .stub(fs, "readJSON")
      .onFirstCall()
      .resolves(pluginManifest)
      .onSecondCall()
      .resolves(manifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
  });

  it("should throw error if failed to run kiota remove command", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId",
          file: "mockedFile",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox
      .stub(mockedDriverContext.ui, "runCommand")
      .onFirstCall()
      .resolves(ok("mockedCommandResult"))
      .onSecondCall()
      .resolves(ok("mockedCommandResult"))
      .onThirdCall()
      .returns(err(new SystemError("mockedSource", "mockedError", "mockedErrorMessage")));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox
      .stub(fs, "readJSON")
      .onFirstCall()
      .resolves(pluginManifest)
      .onSecondCall()
      .resolves(manifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
  });

  it("shoult throw error if no openapi spec generated", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId",
          file: "mockedFile",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox.stub(mockedDriverContext.ui, "runCommand").resolves(ok("mockedCommandResult"));
    sandbox.stub(fs, "readdirSync").returns([] as any);
    sandbox.stub(fs, "readJSON").resolves(pluginManifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.be.equal("NoSpecError");
    }
  });

  it("shoult throw error if no action in da manifest", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox.stub(mockedDriverContext.ui, "runCommand").resolves(ok("mockedCommandResult"));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox.stub(fs, "readJSON").resolves(pluginManifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.be.equal("NoActionError");
    }
  });

  it("shoult throw error if action number = 0 in da manifest", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox.stub(mockedDriverContext.ui, "runCommand").resolves(ok("mockedCommandResult"));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox.stub(fs, "readJSON").resolves(pluginManifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.be.equal("NoActionError");
    }
  });

  it("shoult throw error if action number > 1 in da manifest", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId1",
          file: "mockedFile1",
        },
        {
          id: "mockedActionId2",
          file: "mockedFile2",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox.stub(mockedDriverContext.ui, "runCommand").resolves(ok("mockedCommandResult"));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox.stub(fs, "readJSON").resolves(pluginManifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.be.equal("MultipleActionError");
    }
  });

  it("unhandled error should be handled", async () => {
    const args: TypeSpecCompileArgs = {
      path: "mockedPath",
      manifestPath: "mockedManifestPath",
    };
    const pluginManifest: DeclarativeCopilotManifestSchema = {
      id: "mockedId",
      name: "mockedName",
      description: "mockedDescription",
      actions: [
        {
          id: "mockedActionId",
          file: "mockedFile",
        },
      ],
    };

    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "rmSync").returns();
    sandbox.stub(mockedDriverContext.ui, "runCommand").throws(new Error("mockedError"));
    sandbox.stub(fs, "readdirSync").returns(["openapi.yaml"] as any);
    sandbox
      .stub(fs, "readJSON")
      .onFirstCall()
      .resolves(pluginManifest)
      .onSecondCall()
      .resolves(manifest);
    sandbox.stub(fs, "writeJSON").callsFake((path: string, data: any) => {
      const dataToWrite = JSON.stringify(data);
      expect(dataToWrite.includes("declarativeAgent.json")).to.be.true;
    });
    const result = await typeSpecCompileDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    if (result.result.isErr()) {
      expect(result.result.error.name).to.be.equal("UnhandledError");
    }
  });
});
