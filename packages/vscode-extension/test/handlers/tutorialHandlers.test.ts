import * as sinon from "sinon";
import * as chai from "chai";
import * as globalVariables from "../../src/globalVariables";
import * as localizeUtils from "../../src/utils/localizeUtils";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { OptionItem, err, ok } from "@microsoft/teamsfx-api";
import { TreatmentVariableValue } from "../../src/exp/treatmentVariables";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { openTutorialHandler, selectTutorialsHandler } from "../../src/handlers/tutorialHandlers";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import { PanelType } from "../../src/controls/PanelType";

describe("tutorialHandlers", () => {
  describe("selectTutorialsHandler()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy Path", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sandbox.stub(globalVariables, "isSPFxProject").value(false);
      let tutorialOptions: OptionItem[] = [];
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: (options: any) => {
          tutorialOptions = options.options;
          return Promise.resolve(ok({ type: "success", result: { id: "test", data: "data" } }));
        },
        openUrl: () => Promise.resolve(ok(true)),
      });

      const result = await selectTutorialsHandler();

      chai.assert.equal(tutorialOptions.length, 17);
      chai.assert.isTrue(result.isOk());
      chai.assert.equal(tutorialOptions[1].data, "https://aka.ms/teamsfx-notification-new");
    });

    it("SelectOption returns error", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sandbox.stub(globalVariables, "isSPFxProject").value(false);
      let tutorialOptions: OptionItem[] = [];
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: (options: any) => {
          tutorialOptions = options.options;
          return Promise.resolve(err("error"));
        },
        openUrl: () => Promise.resolve(ok(true)),
      });

      const result = await selectTutorialsHandler();

      chai.assert.equal(tutorialOptions.length, 17);
      chai.assert.equal(result.isErr() ? result.error : "", "error");
    });

    it("SPFx projects - v3", async () => {
      sandbox.stub(localizeUtils, "localize").returns("");
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sandbox.stub(globalVariables, "isSPFxProject").value(true);
      let tutorialOptions: OptionItem[] = [];
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        selectOption: (options: any) => {
          tutorialOptions = options.options;
          return Promise.resolve(ok({ type: "success", result: { id: "test", data: "data" } }));
        },
        openUrl: () => Promise.resolve(ok(true)),
      });

      const result = await selectTutorialsHandler();

      chai.assert.equal(tutorialOptions.length, 1);
      chai.assert.isTrue(result.isOk());
      chai.assert.equal(tutorialOptions[0].data, "https://aka.ms/teamsfx-add-cicd-new");
    });
  });

  describe("openTutorialHandler()", () => {
    const sandbox = sinon.createSandbox();

    afterEach(() => {
      sandbox.restore();
    });

    it("Happy Path", async () => {
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        openUrl: () => Promise.resolve(ok(true)),
      });
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(true);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        openUrl: (link: string) => Promise.resolve(ok(true)),
      });
      const createOrShowStub = sandbox.stub(WebviewPanel, "createOrShow");

      const result = await openTutorialHandler([
        TelemetryTriggerFrom.Auto,
        { id: "cardActionResponse", data: "cardActionResponse" } as OptionItem,
      ]);

      chai.assert.isTrue(result.isOk());
      chai.assert.equal(result.isOk() ? result.value : "Not Equal", undefined);
      chai.assert.isTrue(createOrShowStub.calledOnceWithExactly(PanelType.RespondToCardActions));
    });

    it("Args less than 2", async () => {
      const result = await openTutorialHandler();
      chai.assert.isTrue(result.isOk());
      chai.assert.equal(result.isOk() ? result.value : "Not Equal", undefined);
    });
  });
});
