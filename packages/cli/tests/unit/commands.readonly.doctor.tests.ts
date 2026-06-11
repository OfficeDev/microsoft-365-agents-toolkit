// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLIContext, err, ok, signedIn, signedOut } from "@microsoft/teamsfx-api";
import {
  FuncToolChecker,
  FxCore,
  LocalCertificateManager,
  LtsNodeChecker,
  UserCancelError,
} from "@microsoft/teamsfx-core";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";
import { assert } from "chai";
import * as sinon from "sinon";
import { setCommand } from "../../src/commands/models/set";
import { setSensitivityLabelCommand } from "../../src/commands/models/setSensitivityLabel";
import { DoctorChecker, teamsappDoctorCommand } from "../../src/commands/models/teamsapp/doctor";
import M365TokenProvider from "../../src/commonlib/M365TokenProviderWrapper";

describe("CLI read-only commands doctor", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(process.stdout, "write").returns(true as any);
    sandbox.stub(process.stderr, "write").returns(true as any);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("doctor", async () => {
    describe("checkAccount", async () => {
      it("checkAccount error", async () => {
        sandbox
          .stub(DoctorChecker.prototype, "checkM365Account")
          .resolves(err(new UserCancelError()));
        const checker = new DoctorChecker();
        await checker.checkAccount();
      });
      it("checkAccount success", async () => {
        sandbox.stub(DoctorChecker.prototype, "checkM365Account").resolves(ok("success"));
        const checker = new DoctorChecker();
        await checker.checkAccount();
      });
    });
    describe("checkM365Account", async () => {
      it("checkM365Account - signin", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        sandbox.stub(M365TokenProvider, "getStatus").returns(
          Promise.resolve(
            ok({
              status: signedIn,
              token: token,
              accountInfo: {
                tid: tenantId,
                upn: upn,
              },
            })
          )
        );
        sandbox.stub(DoctorChecker.prototype as any, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "is signed in and custom app upload permission is enabled");
      });
      it("checkM365Account - error", async () => {
        sandbox.stub(M365TokenProvider, "getStatus").resolves(err(new UserCancelError()));
        sandbox.stub(DoctorChecker.prototype as any, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "You've not signed into your Microsoft 365 account yet.");
      });
      it("checkM365Account - error2", async () => {
        sandbox.stub(M365TokenProvider, "getStatus").rejects(new Error("test"));
        sandbox.stub(DoctorChecker.prototype as any, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isErr());
      });
      it("checkM365Account - signout", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        const getStatusStub = sandbox.stub(M365TokenProvider, "getStatus");
        getStatusStub.onCall(0).resolves(
          ok({
            status: signedOut,
          })
        );
        getStatusStub.onCall(1).resolves(
          ok({
            status: signedIn,
            token: token,
            accountInfo: {
              tid: tenantId,
              upn: upn,
            },
          })
        );
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok(token));
        sandbox.stub(DoctorChecker.prototype as any, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "is signed in and custom app upload permission is enabled");
      });

      it("checkM365Account - no custom app upload permission", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        sandbox.stub(M365TokenProvider, "getStatus").returns(
          Promise.resolve(
            ok({
              status: signedIn,
              token: token,
              accountInfo: {
                tid: tenantId,
                upn: upn,
              },
            })
          )
        );
        sandbox.stub(DoctorChecker.prototype as any, "getSideloadingStatus").resolves(false);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const value = (accountRes as any).value;
        assert.include(
          value,
          "Your Microsoft 365 tenant admin hasn't enabled custom app upload permission for your account"
        );
      });
    });

    describe("checkNodejs", async () => {
      it("installed", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: true } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
      it("error", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: true, error: new UserCancelError() } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
      it("not installed", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: false } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
    });
    describe("checkFuncCoreTool", async () => {
      it("installed", async () => {
        sandbox
          .stub(FuncToolChecker.prototype, "queryFuncVersion")
          .resolves({ versionStr: "3.0" } as any);
        const checker = new DoctorChecker();
        await checker.checkFuncCoreTool();
      });
      it("not installed", async () => {
        sandbox.stub(FuncToolChecker.prototype, "queryFuncVersion").rejects(new Error());
        const checker = new DoctorChecker();
        await checker.checkFuncCoreTool();
      });
    });
    describe("checkCert", async () => {
      it("not found", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: false } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
      it("found trusted", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: true, alreadyTrusted: true } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
      it("found not trusted", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: true, alreadyTrusted: false } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
    });

    it("getSideloadingStatus defaults to false when dependency returns undefined", async () => {
      sandbox.stub(tools, "getSideloadingStatus").resolves(undefined);
      const checker = new DoctorChecker();
      const result = await (checker as any).getSideloadingStatus("token");
      assert.isFalse(result);
    });

    it("happy", async () => {
      sandbox.stub(DoctorChecker.prototype, "checkAccount").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkNodejs").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkFuncCoreTool").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkCert").resolves();
      const ctx: CLIContext = {
        command: {
          ...teamsappDoctorCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} doctor`,
        },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappDoctorCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    describe("getSetCommand", async () => {
      it("set command", async () => {
        const commands = setCommand();
        assert.isTrue(commands.commands?.length === 1);
      });
    });

    describe("set sensitivity label", async () => {
      it("success", async () => {
        sandbox.stub(FxCore.prototype, "setSensitivityLabel").resolves(ok(undefined));
        const ctx: CLIContext = {
          command: { ...setSensitivityLabelCommand, fullName: "set sensitivity label" },
          optionValues: {},
          globalOptionValues: {},
          argumentValues: [],
          telemetryProperties: {},
        };
        const res = await setSensitivityLabelCommand.handler!(ctx);
        assert.isTrue(res.isOk());
      });
    });
  });
});
