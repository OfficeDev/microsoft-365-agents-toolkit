import { settings } from "@microsoft/teams-js";
import { settings as settingsAlias } from "@microsoft/teams-js";

interface ITestContext extends settings.SaveEvent {
  x: number;
}

class TestContext implements settings.SaveEvent {
  result: settings.SaveParameters;
  notifySuccess(): void {
    throw new Error("Method not implemented.");
  }
  notifyFailure(reason?: string): void {
    throw new Error("Method not implemented.");
  }
}

interface ITestContextAlias extends settingsAlias.SaveEvent {
  x: number;
}

class TestContextAlias implements settingsAlias.SaveEvent {
  result: settingsAlias.SaveParameters;
  notifySuccess(): void {
    throw new Error("Method not implemented.");
  }
  notifyFailure(reason?: string): void {
    throw new Error("Method not implemented.");
  }
}
