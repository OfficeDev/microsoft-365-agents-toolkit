// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as sinon from "sinon";
import chai from "chai";
import fs from "fs-extra";
import { copyAppPackageToSPFxDriver } from "../../../../src/component/driver/teamsApp/copyAppPackageToSPFx";
import { copyAppPackageToSPFxArgs } from "../../../../src/component/driver/teamsApp/interfaces/CopyAppPackageToSPFxArgs";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import chaiAsPromised from "chai-as-promised";
import AdmZip from "adm-zip";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("teamsApp/copyAppPackageToSPFx", async () => {
  const driver = new copyAppPackageToSPFxDriver();
  const args: copyAppPackageToSPFxArgs = {
    appPackagePath: "./teamsApp/a.zip",
    spfxFolder: "./SPFx",
  };
  const mockedDriverContext: any = { projectPath: "C://TeamsApp" };

  afterEach(() => {
    sinon.restore();
  });

  it("should successfully copy app package for SPFx", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "copyFile");
    sinon.stub(fs, "writeFile");
    sinon.stub(fs, "readdir").resolves(["color.png", "outline.png"] as any);
    sinon
      .stub(copyAppPackageToSPFxDriver.prototype, "getIcons")
      .resolves({ color: Buffer.from("color.png"), outline: Buffer.from("outline.png") });

    const result = await driver.execute(args, mockedDriverContext);
    expect(result.result.isOk()).to.be.true;
    expect(result.summaries.length).to.eq(2);
  });

  it("fail to copy app package for SPFx - FileNotFoundError", async () => {
    sinon.stub(fs, "pathExists").resolves(false);

    const result = await driver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect((result.result as any).error.name).to.be.equal(AppStudioError.FileNotFoundError.name);
  });

  it("should successfully get icons", async () => {
    const zip = new AdmZip();
    zip.addFile(
      Constants.MANIFEST_FILE,
      Buffer.from(JSON.stringify({ icons: { color: "color.png", outline: "outline.png" } }))
    );
    zip.addFile("./resources/color.png", Buffer.from(""));
    zip.addFile("./resources/outline.png", Buffer.from(""));
    sinon.stub(fs, "readFile").resolves(zip.toBuffer());
    await expect(driver.getIcons(args.appPackagePath)).to.eventually.deep.equal({
      color: Buffer.from(""),
      outline: Buffer.from(""),
    });
  });

  it("fail to get icons - FileNotFoundError", async () => {
    sinon.stub(fs, "readFile").resolves(undefined);
    await expect(driver.getIcons(args.appPackagePath)).to.be.rejectedWith();
  });
});
