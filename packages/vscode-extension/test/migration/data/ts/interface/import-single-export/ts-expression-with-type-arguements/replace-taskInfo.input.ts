import { TaskInfo } from "@microsoft/teams-js";
import { TaskInfo as TaskInfoAlias } from "@microsoft/teams-js";

interface ITestContext extends TaskInfo {
  x: number;
}

class TestContext implements TaskInfo {}

interface ITestContextAlias extends TaskInfoAlias {
  x: number;
}

class TestContextAlias implements TaskInfoAlias {}
