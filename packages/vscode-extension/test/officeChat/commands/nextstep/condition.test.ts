import * as chai from "chai";
import {
  canOfficeAddInPreviewInLocalEnv,
  isDebugSucceededAfterSourceCodeChanged,
  isDependenciesInstalled,
  isDidNoActionAfterScaffolded,
  isHaveReadMe,
  isProjectOpened,
  isNodeInstalled,
} from "../../../../src/officeChat/commands/nextStep/condition";
import { OfficeWholeStatus } from "../../../../src/officeChat/commands/nextStep/types";
import { emptyProjectStatus } from "../../../../src/utils/projectStatusUtils";
import { CommandKey } from "../../../../src/constants";

describe("offce chat nextstep conditions", () => {
  it("isProjectOpened", () => {
    chai.assert.isTrue(
      isProjectOpened({
        projectOpened: {},
      } as OfficeWholeStatus)
    );
    chai.assert.isFalse(isProjectOpened({} as OfficeWholeStatus));
  });

  describe("isNodeInstalled", () => {
    it("isNodeInstalled", () => {
      chai.assert.isTrue(
        isNodeInstalled({
          projectOpened: {
            isNodeInstalled: true,
          },
          machineStatus: {},
        } as OfficeWholeStatus)
      );

      chai.assert.isFalse(
        isNodeInstalled({
          projectOpened: {
            isNodeInstalled: false,
          },
          machineStatus: {},
        } as OfficeWholeStatus)
      );

      chai.assert.isFalse(isNodeInstalled({} as OfficeWholeStatus));
    });
  });

  describe("isDidNoActionAfterScaffolded", () => {
    it("no opened project", () => {
      chai.assert.isTrue(isDidNoActionAfterScaffolded({} as OfficeWholeStatus));
    });

    it("action status is empty", () => {
      chai.assert.isTrue(
        isDidNoActionAfterScaffolded({
          projectOpened: {
            actionStatus: emptyProjectStatus(),
          },
        } as OfficeWholeStatus)
      );
    });

    it("some action is done", () => {
      chai.assert.isFalse(
        isDidNoActionAfterScaffolded({
          projectOpened: {
            actionStatus: {
              ...emptyProjectStatus(),
              [CommandKey.Provision]: { result: "success", time: new Date() },
            },
          },
        } as OfficeWholeStatus)
      );
    });

    it("some action is failed", () => {
      chai.assert.isFalse(
        isDidNoActionAfterScaffolded({
          projectOpened: {
            actionStatus: {
              ...emptyProjectStatus(),
              [CommandKey.Provision]: { result: "fail", time: new Date() },
            },
          },
        } as OfficeWholeStatus)
      );
    });
  });

  describe("isDebugSucceededAfterSourceCodeChanged", () => {
    it("no opened project", () => {
      chai.assert.isFalse(isDebugSucceededAfterSourceCodeChanged({} as OfficeWholeStatus));
    });

    it("local debug not run before", () => {
      chai.assert.isFalse(
        isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "no run", time: new Date() },
            },
          },
        } as OfficeWholeStatus)
      );
    });

    it("local debug failed before", () => {
      chai.assert.isFalse(
        isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "fail", time: new Date() },
            },
          },
        } as OfficeWholeStatus)
      );
    });

    it("local debug succeeded before but out of date", () => {
      chai.assert.isFalse(
        isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "success", time: new Date(0) },
            },
            codeModifiedTime: {
              source: new Date(),
            },
          },
        } as OfficeWholeStatus)
      );
    });

    it("local debug succeeded after source changed", () => {
      chai.assert.isTrue(
        isDebugSucceededAfterSourceCodeChanged({
          projectOpened: {
            actionStatus: {
              [CommandKey.LocalDebug]: { result: "success", time: new Date() },
            },
            codeModifiedTime: {
              source: new Date(0),
            },
          },
        } as OfficeWholeStatus)
      );
    });
  });

  describe("canOfficeAddInPreviewInLocalEnv", () => {
    it('should return true when launchJSONContent includes "desktop (edge legacy)" or "desktop (edge chromium)"', () => {
      const result = canOfficeAddInPreviewInLocalEnv({
        projectOpened: {
          launchJSONContent: "desktop (edge legacy)",
        },
      } as OfficeWholeStatus);
      chai.assert.isTrue(result);
    });

    it('should return false when launchJSONContent does not include "desktop (edge legacy)" or "desktop (edge chromium)"', () => {
      const result = canOfficeAddInPreviewInLocalEnv({
        projectOpened: {
          launchJSONContent: "",
        },
      } as OfficeWholeStatus);
      chai.assert.isFalse(result);
    });

    it("should return false when projectOpened or launchJSONContent is not defined", () => {
      const result = canOfficeAddInPreviewInLocalEnv({} as OfficeWholeStatus);
      chai.assert.isFalse(result);
    });
  });

  describe("isDependenciesInstalled", () => {
    it("isDependenciesInstalled", () => {
      chai.assert.isTrue(
        isDependenciesInstalled({
          projectOpened: {
            nodeModulesExist: true,
          },
          machineStatus: {},
        } as OfficeWholeStatus)
      );

      chai.assert.isFalse(
        isDependenciesInstalled({
          projectOpened: {
            nodeModulesExist: false,
          },
          machineStatus: {},
        } as OfficeWholeStatus)
      );

      chai.assert.isFalse(isDependenciesInstalled({} as OfficeWholeStatus));
    });
  });

  describe("isHaveReadMe", () => {
    it("no opened project", () => {
      chai.assert.isFalse(isHaveReadMe({} as OfficeWholeStatus));
    });

    it("no readme", () => {
      chai.assert.isFalse(
        isHaveReadMe({
          projectOpened: {},
        } as OfficeWholeStatus)
      );
    });

    it("had readme", () => {
      chai.assert.isTrue(
        isHaveReadMe({
          projectOpened: {
            readmeContent: "123123",
          },
        } as OfficeWholeStatus)
      );
    });
  });
});
