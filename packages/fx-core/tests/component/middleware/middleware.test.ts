import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { performance } from "perf_hooks";
import sinon from "sinon";
import { setTools } from "../../../src/common/globalVars";
import { TelemetryProperty } from "../../../src/common/telemetry";
import { TeamsFxTelemetryReporter } from "../../../src/component/utils/teamsFxTelemetryReporter";
import { MockTools } from "../../core/utils";
import { MockDriver } from "./helper";

chai.use(chaiAsPromised);

describe("Action Middleware", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  setTools(tools);

  afterEach(() => {
    sandbox.restore();
  });

  it("addStartAndEndTelemetry send correct cost time", async () => {
    const perfStub = sandbox.stub(performance, "now");
    perfStub.onFirstCall().returns(0);
    perfStub.onSecondCall().returns(1000);
    sandbox.stub(TeamsFxTelemetryReporter.prototype, "sendStartEvent");
    const sendEndEventStub = sandbox.stub(TeamsFxTelemetryReporter.prototype, "sendEndEvent");
    sendEndEventStub.callsFake((config) => {
      chai.assert.equal(config.measurements?.[TelemetryProperty.TimeCost], 1000);
    });

    await new MockDriver().execute(undefined, { telemetryReporter: {} as any } as any);

    chai.assert.isTrue(sendEndEventStub.called);
  });
});
