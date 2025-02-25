import { OptionItem, err, ok } from "@microsoft/teamsfx-api";
import * as templateMetadata from "@microsoft/teamsfx-core/build/component/generator/templates/metadata";
import * as chai from "chai";
import * as sinon from "sinon";
import { PanelType } from "../../src/controls/PanelType";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import { TreatmentVariableValue } from "../../src/exp/treatmentVariables";
import * as globalVariables from "../../src/globalVariables";
import { openTutorialHandler, selectTutorialsHandler } from "../../src/handlers/tutorialHandlers";
import * as vsc_ui from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";
import * as localizeUtils from "../../src/utils/localizeUtils";

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
      const createOrShowStub = sandbox.stub(WebviewPanel, "createOrShow");

      const result = await openTutorialHandler([
        TelemetryTriggerFrom.Auto,
        { id: "cardActionResponse", data: "cardActionResponse" } as OptionItem,
      ]);

      chai.assert.isTrue(result.isOk());
      chai.assert.equal(result.isOk() ? result.value : "Not Equal", undefined);
      chai.assert.isTrue(createOrShowStub.calledOnceWithExactly(PanelType.RespondToCardActions));
    });

    it("Template option", async () => {
      let openLink = "";
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(TreatmentVariableValue, "inProductDoc").value(false);
      sandbox.stub(vsc_ui, "VS_CODE_UI").value({
        openUrl: (link: string) => {
          openLink = link;
          return Promise.resolve(ok(true));
        },
      });
      sandbox.stub(templateMetadata, "getDefaultTemplatesOnPlatform").returns([
        {
          id: "test",
          description: "test",
          language: "none",
          name: "test",
          link: "testLink",
        },
      ]);

      const result = await openTutorialHandler([
        TelemetryTriggerFrom.Auto,
        { id: "test", data: "test" } as OptionItem,
      ]);

      chai.assert.isTrue(result.isOk());
      chai.assert.equal(openLink, "testLink");
    });

    it("Args less than 2", async () => {
      const result = await openTutorialHandler();
      chai.assert.isTrue(result.isOk());
      chai.assert.equal(result.isOk() ? result.value : "Not Equal", undefined);
    });
  });
});
