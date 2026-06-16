// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Context, err, Inputs, ok, Platform } from "@microsoft/teamsfx-api";
import chai from "chai";
import sinon from "sinon";
import * as uuid from "uuid";

import { setTools } from "../../../../src/common/globalVars";
import { addWebPartDeps, AddWebPartDriver } from "../../../../src/component/driver/add/addWebPart";
import { NoConfigurationError } from "../../../../src/component/driver/add/error/noConfigurationError";
import { AddWebPartArgs } from "../../../../src/component/driver/add/interface/AddWebPartArgs";
import { AppStudioResultFactory } from "../../../../src/component/driver/teamsApp/results";
import { InstallSoftwareError } from "../../../../src/error/common";
import { MockedM365Provider, MockTools } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

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

  beforeEach(() => {
    setTools(new MockTools());
  });

  it("Returns error when no .yo-rc.json file exist", async () => {
    sinon.stub(addWebPartDeps, "pathExists").resolves(false);

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isErr()).to.be.true;
    chai.expect((res as any).error).instanceOf(NoConfigurationError);
  });

  it("Returns error when Yeoman scaffold fails", async () => {
    sinon.stub(addWebPartDeps, "pathExists").resolves(true);
    sinon
      .stub(addWebPartDeps, "doYeomanScaffold")
      .resolves(err(new InstallSoftwareError("spfx", "yo")));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isErr()).to.be.true;
  });

  it("Returns error when updating manifest fails", async () => {
    sinon.stub(addWebPartDeps, "pathExists").resolves(true);
    const componentId = uuid.v4();
    sinon.stub(addWebPartDeps, "doYeomanScaffold").resolves(ok(componentId));
    sinon
      .stub(addWebPartDeps, "addCapabilities")
      .resolves(err(AppStudioResultFactory.UserError("test", ["test msg", "test msg"])));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isErr()).to.be.true;
  });

  it("Returns success when add web part OK", async () => {
    sinon.stub(addWebPartDeps, "pathExists").resolves(true);
    const componentId = uuid.v4();
    const doYeomanScaffoldStub = sinon
      .stub(addWebPartDeps, "doYeomanScaffold")
      .resolves(ok(componentId));
    const addCapabilitiesStub = sinon
      .stub(addWebPartDeps, "addCapabilities")
      .resolves(ok(undefined));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isOk(), res.isErr() ? String(res.error?.message ?? res.error) : undefined).to.be
      .true;
    chai.expect(doYeomanScaffoldStub.calledOnce).to.be.true;
    chai.expect(addCapabilitiesStub.calledTwice).to.be.true;
  });

  it("Returns success when add web part for SPFx higher than 1.21", async () => {
    sinon.stub(addWebPartDeps, "pathExists").resolves(true);
    const componentId = uuid.v4();
    const doYeomanScaffoldStub = sinon
      .stub(addWebPartDeps, "doYeomanScaffold")
      .callsFake(async (SPFxContext: Context, inputs: Inputs, projectPath: string) => {
        if (!SPFxContext.templateVariables) {
          SPFxContext.templateVariables = {};
        }
        SPFxContext.templateVariables!["useNewDevUrl"] = "true";
        return Promise.resolve(ok(componentId));
      });
    const addCapabilitiesStub = sinon
      .stub(addWebPartDeps, "addCapabilities")
      .resolves(ok(undefined));

    const res = await driver.run(args, mockedDriverContext);

    chai.expect(res.isOk(), res.isErr() ? String(res.error?.message ?? res.error) : undefined).to.be
      .true;
    chai.expect(doYeomanScaffoldStub.calledOnce).to.be.true;
    chai.expect(addCapabilitiesStub.calledTwice).to.be.true;
  });
});
