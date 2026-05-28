import { TaskInfo } from "@microsoft/teams-js";
import { TaskInfo as TaskInfoAlias } from "@microsoft/teams-js";

async function f1(x: TaskInfo): Promise<TaskInfo> {
  return x;
}

async function f2(x: TaskInfoAlias): Promise<TaskInfoAlias> {
  return x;
}

const x1: TaskInfo | undefined = undefined;
const x2: TaskInfoAlias | undefined = undefined;

if (!!x1) {
  const y1 = x1 as TaskInfo;
  //const y2 = <TaskInfo>x1;
}
if (!!x2) {
  const y1 = x2 as TaskInfoAlias;
  //const y2 = <TaskInfoAlias>x2;
}

class TestClass {
  private x1: TaskInfo;
  private x2: TaskInfoAlias;

  constructor(x1: TaskInfo, x2: TaskInfoAlias) {
    this.x1 = x1;
    this.x2 = x2;
  }
}

interface ITestInterface {
  x1: TaskInfo;
  x2: TaskInfoAlias;
}
