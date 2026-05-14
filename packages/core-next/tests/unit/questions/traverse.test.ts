/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import { ok, err } from "neverthrow";
import { traverseQuestionTree } from "../../../src/questions/traverse";
import type { IQTreeNode, Group } from "../../../src/api/qm/question";
import type { InputResult } from "../../../src/api/qm/ui";
import type { Inputs, OptionItem } from "../../../src/api/types";
import { Platform } from "../../../src/api/constants";
import { createMockContext } from "../testHelper";

function makeInputs(overrides?: Record<string, unknown>): Inputs {
  return { platform: Platform.VSCode, ...overrides };
}

function groupNode(children?: IQTreeNode[]): IQTreeNode {
  return {
    data: { type: "group" } as Group,
    children,
  };
}

function selectNode(
  name: string,
  options: (string | OptionItem)[],
  opts?: {
    children?: IQTreeNode[];
    condition?: IQTreeNode["condition"];
    skipSingleOption?: boolean;
  }
): IQTreeNode {
  return {
    data: {
      type: "singleSelect",
      name,
      title: `Select ${name}`,
      staticOptions: options,
      skipSingleOption: opts?.skipSingleOption,
    } as any,
    condition: opts?.condition,
    children: opts?.children,
  };
}

function textNode(
  name: string,
  opts?: { condition?: IQTreeNode["condition"]; defaultVal?: string }
): IQTreeNode {
  return {
    data: {
      type: "text",
      name,
      title: `Enter ${name}`,
      default: opts?.defaultVal,
    } as any,
    condition: opts?.condition,
  };
}

function folderNode(name: string): IQTreeNode {
  return {
    data: {
      type: "folder",
      name,
      title: `Select ${name}`,
    } as any,
  };
}

describe("traverseQuestionTree", () => {
  let sandbox: sinon.SinonSandbox;
  let ctx: ReturnType<typeof createMockContext>;

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    ctx = createMockContext();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should traverse a single question and store the answer", async () => {
    const tree = groupNode([selectNode("color", ["red", "blue", "green"])]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "blue" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["color"]).to.equal("blue");
  });

  it("should traverse multiple sequential questions", async () => {
    const tree = groupNode([
      selectNode("projectType", ["bot", "tab"]),
      textNode("appName"),
      folderNode("folder"),
    ]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "bot" } as InputResult<string>)
    );
    (ctx.ui.inputText as sinon.SinonStub).resolves(
      ok({ type: "success", result: "my-app" } as InputResult<string>)
    );
    (ctx.ui.selectFolder as sinon.SinonStub).resolves(
      ok({ type: "success", result: "/home/user/projects" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["projectType"]).to.equal("bot");
    expect(inputs["appName"]).to.equal("my-app");
    expect(inputs["folder"]).to.equal("/home/user/projects");
  });

  it("should skip questions when inputs are pre-filled", async () => {
    const tree = groupNode([selectNode("color", ["red", "blue"]), textNode("name")]);
    const inputs = makeInputs({ color: "red" });

    (ctx.ui.inputText as sinon.SinonStub).resolves(
      ok({ type: "success", result: "test-name" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["color"]).to.equal("red"); // kept pre-filled
    expect(inputs["name"]).to.equal("test-name");
    expect((ctx.ui.selectOption as sinon.SinonStub).callCount).to.equal(0); // never asked
  });

  it("should evaluate StringValidation equals condition", async () => {
    const tree = groupNode([
      selectNode("projectType", ["bot", "tab"], {
        children: [
          textNode("botName", { condition: { equals: "bot" } }),
          textNode("tabName", { condition: { equals: "tab" } }),
        ],
      }),
    ]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "bot" } as InputResult<string>)
    );
    (ctx.ui.inputText as sinon.SinonStub).resolves(
      ok({ type: "success", result: "my-bot" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["botName"]).to.equal("my-bot");
    expect(inputs["tabName"]).to.be.undefined; // condition not met
  });

  it("should evaluate ConditionFunc condition", async () => {
    const tree = groupNode([
      selectNode("mode", ["basic", "advanced"]),
      textNode("advancedOption", {
        condition: ((inputs: Inputs) => inputs["mode"] === "advanced") as any,
      }),
    ]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "basic" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["advancedOption"]).to.be.undefined; // condition not met
  });

  it("should evaluate enum condition", async () => {
    const tree = groupNode([
      selectNode("lang", ["ts", "js", "py"], {
        children: [textNode("tsConfig", { condition: { enum: ["ts", "js"] } })],
      }),
    ]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "py" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["tsConfig"]).to.be.undefined; // py not in [ts, js]
  });

  it("should pass enum condition when value matches", async () => {
    const tree = groupNode([
      selectNode("lang", ["ts", "js", "py"], {
        children: [textNode("tsConfig", { condition: { enum: ["ts", "js"] } })],
      }),
    ]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "ts" } as InputResult<string>)
    );
    (ctx.ui.inputText as sinon.SinonStub).resolves(
      ok({ type: "success", result: "strict" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["tsConfig"]).to.equal("strict");
  });

  it("should handle back navigation", async () => {
    const tree = groupNode([selectNode("step1", ["a", "b"]), selectNode("step2", ["x", "y"])]);
    const inputs = makeInputs();

    const selectStub = ctx.ui.selectOption as sinon.SinonStub;
    // First call: answer step1 with "a"
    selectStub.onFirstCall().resolves(ok({ type: "success", result: "a" } as InputResult<string>));
    // Second call: go back
    selectStub.onSecondCall().resolves(ok({ type: "back" } as InputResult<string>));
    // Third call (re-asks step1): answer with "b"
    selectStub.onThirdCall().resolves(ok({ type: "success", result: "b" } as InputResult<string>));
    // Fourth call (step2): answer "y"
    selectStub.onCall(3).resolves(ok({ type: "success", result: "y" } as InputResult<string>));

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["step1"]).to.equal("b");
    expect(inputs["step2"]).to.equal("y");
  });

  it("should return error when backing past the first question", async () => {
    const tree = groupNode([selectNode("q1", ["a", "b"])]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(ok({ type: "back" } as InputResult<string>));

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("UserCancelled");
    }
  });

  it("should return error when UI returns an error", async () => {
    const tree = groupNode([textNode("name")]);
    const inputs = makeInputs();

    (ctx.ui.inputText as sinon.SinonStub).resolves(
      err({ name: "FxError", message: "Connection lost" } as any)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("QuestionTraversalError");
    }
  });

  it("should extract id from OptionItem", async () => {
    const tree = groupNode([
      selectNode("template", [
        { id: "bot-basic", label: "Basic Bot" },
        { id: "tab-react", label: "React Tab" },
      ]),
    ]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({
        type: "success",
        result: { id: "tab-react", label: "React Tab" },
      } as InputResult<OptionItem>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["template"]).to.equal("tab-react");
  });

  it("should auto-skip singleSelect with one option and skipSingleOption=true", async () => {
    const tree = groupNode([selectNode("lang", ["typescript"], { skipSingleOption: true })]);
    const inputs = makeInputs();

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["lang"]).to.equal("typescript");
    expect((ctx.ui.selectOption as sinon.SinonStub).callCount).to.equal(0);
  });

  it("should handle an empty tree", async () => {
    const tree = groupNode();
    const inputs = makeInputs();

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
  });

  it("should skip nodes with inputsDisabled=self", async () => {
    const node = textNode("disabled-q");
    node.inputsDisabled = "self";
    const tree = groupNode([node, textNode("enabled-q")]);
    const inputs = makeInputs();

    (ctx.ui.inputText as sinon.SinonStub).resolves(
      ok({ type: "success", result: "val" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["disabled-q"]).to.be.undefined;
    expect(inputs["enabled-q"]).to.equal("val");
  });

  it("should handle nested group-then-question structure", async () => {
    const tree = groupNode([groupNode([selectNode("nested-q", ["opt1", "opt2"])])]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "opt1" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["nested-q"]).to.equal("opt1");
  });

  it("should handle contains condition on array values", async () => {
    const tree = groupNode([
      textNode("features-detail", {
        condition: { contains: "sso" } as any,
      }),
    ]);
    // Pre-fill parent-like value (but since this is a root-level condition with no parent,
    // the condition evaluator checks parentValue which is undefined → condition fails)
    const inputs = makeInputs();

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    // No parent question → contains condition with undefined parent → skipped
    expect(inputs["features-detail"]).to.be.undefined;
  });

  it("should evaluate contains condition with parent question", async () => {
    // Parent is a multiSelect asking for features
    const featuresNode: IQTreeNode = {
      data: {
        type: "multiSelect",
        name: "features",
        title: "Select features",
        staticOptions: ["sso", "api", "db"],
      } as any,
      children: [textNode("sso-config", { condition: { contains: "sso" } })],
    };
    const tree = groupNode([featuresNode]);
    const inputs = makeInputs();

    (ctx.ui.selectOptions as sinon.SinonStub).resolves(
      ok({ type: "success", result: ["sso", "api"] } as InputResult<string[]>)
    );
    (ctx.ui.inputText as sinon.SinonStub).resolves(
      ok({ type: "success", result: "my-sso-config" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["sso-config"]).to.equal("my-sso-config");
  });

  it("should handle confirm question type", async () => {
    const confirmNode: IQTreeNode = {
      data: {
        type: "confirm",
        name: "proceed",
        title: "Continue?",
      } as any,
    };
    const tree = groupNode([confirmNode]);
    const inputs = makeInputs();

    (ctx.ui.confirm as sinon.SinonStub).resolves(
      ok({ type: "success", result: true } as InputResult<boolean>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["proceed"]).to.equal(true);
  });

  it("should handle folder question type", async () => {
    const tree = groupNode([folderNode("destination")]);
    const inputs = makeInputs();

    (ctx.ui.selectFolder as sinon.SinonStub).resolves(
      ok({ type: "success", result: "/home/user/dev" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["destination"]).to.equal("/home/user/dev");
  });

  it("should handle dynamicOptions for singleSelect", async () => {
    const dynamicNode: IQTreeNode = {
      data: {
        type: "singleSelect",
        name: "dynamic",
        title: "Pick one",
        staticOptions: [],
        dynamicOptions: async () => ["alpha", "beta", "gamma"],
      } as any,
    };
    const tree = groupNode([dynamicNode]);
    const inputs = makeInputs();

    (ctx.ui.selectOption as sinon.SinonStub).resolves(
      ok({ type: "success", result: "gamma" } as InputResult<string>)
    );

    const result = await traverseQuestionTree(tree, ctx.ui, inputs);
    expect(result.isOk()).to.be.true;
    expect(inputs["dynamic"]).to.equal("gamma");
  });
});
