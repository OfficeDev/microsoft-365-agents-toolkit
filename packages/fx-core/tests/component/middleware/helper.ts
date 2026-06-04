import { hooks } from "@feathersjs/hooks/lib";
import { FxError, ok, Result } from "@microsoft/teamsfx-api";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import { StepDriver, ExecutionResult } from "../../../src/component/driver/interface/stepDriver";
import { addStartAndEndTelemetry } from "../../../src/component/driver/middleware/addStartAndEndTelemetry";

export class MockDriver implements StepDriver {
  @hooks([addStartAndEndTelemetry("mock", "mock")])
  async execute(args: unknown, context: DriverContext): Promise<ExecutionResult> {
    return { result: ok(new Map<string, string>()), summaries: [] };
  }
}
