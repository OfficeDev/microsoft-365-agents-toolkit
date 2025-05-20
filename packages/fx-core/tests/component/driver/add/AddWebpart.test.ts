// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import { Context, err, Inputs, ok, Platform } from "@microsoft/teamsfx-api";
import chai from "chai";
import fs from "fs-extra";
import sinon from "sinon";
import * as path from "path";
import * as uuid from "uuid";

import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";
import { AddWebPartDriver } from "../../../../src/component/driver/add/addWebPart";
import { AddWebPartArgs } from "../../../../src/component/driver/add/interface/AddWebPartArgs";
import { Constants } from "../../../../src/component/driver/add/utility/constants";
import { NoConfigurationError } from "../../../../src/component/driver/add/error/noConfigurationError";
import { SPFxGenerator } from "../../../../src/component/generator/spfx/spfxGenerator";
import { ManifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { AppStudioResultFactory } from "../../../../src/component/driver/teamsApp/results";
import { setTools } from "../../../../src/common/globalVars";
import { InstallSoftwareError } from "../../../../src/error/common";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";

describe("Add web part driver", async () => {
  const args: AddWebPartArgs = {
    spfxFolder: "C://TeamsApp//src",
    webpartName: "HelloWorld",
    manifestPath: "C://TeamsApp//appPackage//manifest.json",
    localManifestPath: "C://TeamsApp//appPackage//manifest.local.json",
    spfxPackage: "installLocally",
  };
  const driver = new AddWebPartDriver();
  const mockedDriverContext: any = {
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    m365TokenProvider: new MockedM365Provider(),
    platform: Platform.VSCode,
    projectPath: "C://TeamsApp",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("Returns error when no .yo-rc.json file exist", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (directory) => {
      if (directory === path.join(args.spfxFolder, Constants.YO_RC_FILE)) {
        return false;
      }
    });

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isErr()).to.be.true;
    chai.expect((res as any).error).instanceOf(NoConfigurationError);
  });

  it("Returns error when Yeoman scaffold fails", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (directory) => {
      if (directory === path.join(args.spfxFolder, Constants.YO_RC_FILE)) {
        return true;
      }
    });
    sinon
      .stub(SPFxGenerator, "doYeomanScaffold")
      .resolves(err(new InstallSoftwareError("spfx", "yo")));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isErr()).to.be.true;
  });

  it("Returns error when updating manifest fails", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (directory) => {
      if (directory === path.join(args.spfxFolder, Constants.YO_RC_FILE)) {
        return true;
      }
    });
    const componentId = uuid.v4();
    sinon.stub(SPFxGenerator, "doYeomanScaffold").resolves(ok(componentId));
    sinon
      .stub(ManifestUtils.prototype, "addCapabilities")
      .resolves(err(AppStudioResultFactory.UserError("test", ["test msg", "test msg"])));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isErr()).to.be.true;
  });

  it("Returns success when add web part OK", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (directory) => {
      if (directory === path.join(args.spfxFolder, Constants.YO_RC_FILE)) {
        return true;
      }
    });
    const componentId = uuid.v4();
    setTools({
      logProvider: new MockedLogProvider(),
      ui: new MockedUserInteraction(),
      tokenProvider: {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      },
    });
    const doYeomanScaffoldStub = sinon
      .stub(SPFxGenerator, "doYeomanScaffold")
      .resolves(ok(componentId));
    const addCapabilitiesStub = sinon
      .stub(ManifestUtils.prototype, "addCapabilities")
      .resolves(ok(undefined));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isOk()).to.be.true;
    chai.expect(doYeomanScaffoldStub.calledOnce).to.be.true;
    chai.expect(addCapabilitiesStub.calledTwice).to.be.true;
  });

  it("Returns success when add web part for SPFx higher than 1.21", async () => {
    sinon.stub(fs, "pathExists").callsFake(async (directory) => {
      if (directory === path.join(args.spfxFolder, Constants.YO_RC_FILE)) {
        return true;
      }
    });
    const componentId = uuid.v4();
    setTools({
      logProvider: new MockedLogProvider(),
      ui: new MockedUserInteraction(),
      tokenProvider: {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      },
    });
    const doYeomanScaffoldStub = sinon
      .stub(SPFxGenerator, "doYeomanScaffold")
      .callsFake(async (SPFxContext: Context, inputs: Inputs, projectPath: string) => {
        if (!SPFxContext.templateVariables) {
          SPFxContext.templateVariables = {};
        }
        SPFxContext.templateVariables!["useNewDevUrl"] = "true";
        return Promise.resolve(ok(componentId));
      });
    const addCapabilitiesStub = sinon
      .stub(ManifestUtils.prototype, "addCapabilities")
      .resolves(ok(undefined));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isOk()).to.be.true;
    chai.expect(doYeomanScaffoldStub.calledOnce).to.be.true;
    chai.expect(addCapabilitiesStub.calledTwice).to.be.true;
  });
});
