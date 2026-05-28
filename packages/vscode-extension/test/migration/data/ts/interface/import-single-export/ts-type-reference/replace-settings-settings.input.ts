import { settings } from "@microsoft/teams-js";
import { settings as settingsAlias } from "@microsoft/teams-js";

async function f1(x: settings.Settings): Promise<settings.Settings> {
  return x;
}

async function f2(x: settingsAlias.Settings): Promise<settingsAlias.Settings> {
  return x;
}

const x1: settings.Settings | undefined = undefined;
const x2: settingsAlias.Settings | undefined = undefined;

if (!!x1) {
  const y1 = x1 as settings.Settings;
  //const y2 = <settings.Settings>x1;
}
if (!!x2) {
  const y1 = x2 as settingsAlias.Settings;
  //const y2 = <settingsAlias.Settings>x2;
}

class TestClass {
  private x1: settings.Settings;
  private x2: settingsAlias.Settings;

  constructor(x1: settings.Settings, x2: settingsAlias.Settings) {
    this.x1 = x1;
    this.x2 = x2;
  }
}

interface ITestInterface {
  x1: settings.Settings;
  x2: settingsAlias.Settings;
}
