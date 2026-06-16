import * as chai from "chai";
import { LocalDebugPorts, resetLocalDebugPorts } from "../../src/globalVariables";
import { vi } from "vitest";

describe("GlobalVariables", () => {
  it("resetLocalDebugPorts", async () => {
    resetLocalDebugPorts();
    chai.assert.deepEqual(LocalDebugPorts, {
      checkPorts: [],
      conflictPorts: [],
      terminateButton: "",
      process2conflictPorts: {},
      terminateProcesses: [],
    });
  });
});
