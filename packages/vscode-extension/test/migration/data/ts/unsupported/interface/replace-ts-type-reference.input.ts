import { Context } from "@microsoft/teams-js";
import { Context as ContextAlias } from "@microsoft/teams-js";

async function f1(x: Context): Promise<Context> {
  return x;
}

async function f2(x: ContextAlias): Promise<ContextAlias> {
  return x;
}

const x1: Context | undefined = undefined;
const x2: ContextAlias | undefined = undefined;

if (!!x1) {
  const y1 = x1 as Context;
  const y2 = <Context>x1;
}
if (!!x2) {
  const y1 = x2 as ContextAlias;
  const y2 = <ContextAlias>x2;
}

class TestClass {
  private x1: Context;
  private x2: ContextAlias;

  constructor(x1: Context, x2: ContextAlias) {
    this.x1 = x1;
    this.x2 = x2;
  }
}

interface ITestInterface {
  x1: Context;
  x2: ContextAlias;
}
