// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 *
 * This is wrapper of mocha.it which help to inject meta data
 * to mocha.Context
 *
 * Usage:
 *  import "it" in this file instead of mocha.it
 *
 * @example
 * Here's an example with only title
 * ```
 * it("sample: only title");
 *
 * it("should run as normal mocha.it with sync function", function () {
 *   expect(1).equals(1);
 * });
 *
 * it("should run as normal mocha.it with async function", async function () {
 *   expect(1).equals(1);
 * });
 *
 * it("should run as normal mocha.it with sync arrow function", () => {
 *   expect(1).equals(1);
 * });
 *
 * it("should run as normal mocha.it with async arrow function", async () => {
 *   expect(1).equals(1);
 * });
 *
 * it("should inject ctx with sync function", { testPlanCaseId: 1 }, function () {
 *   expect(1).equals(1);
 * });
 *
 * it("should inject ctx with async function", { testPlanCaseId: 1 }, async function () {
 *   expect(1).equals(1);
 * });
 *
 * it("should inject ctx with sync arrow function", { testPlanCaseId: 1 }, () => {
 *   expect(1).equals(1);
 * });
 *
 * it("should inject ctx with async arrow function", { testPlanCaseId: 1 }, async () => {
 *   expect(1).equals(1);
 * });
 * ```
 *
 */
import addContext from "mochawesome/addContext";
import mocha from "mocha";

export function it(title: string, fn?: mocha.AsyncFunc): mocha.Test;
export function it(title: string, fn?: mocha.Func): mocha.Test;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function it(
  title: string,
  ctx: Record<string, any>,
  fn: mocha.Func
): mocha.Test;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function it(
  title: string,
  ctx: Record<string, any>,
  fn: mocha.AsyncFunc
): mocha.Test;

export function it(
  title: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ctxOrFn?: Record<string, any> | mocha.Func | mocha.AsyncFunc,
  fn?: mocha.Func | mocha.AsyncFunc
): mocha.Test {
  let t: mocha.Test;

  if (ctxOrFn) {
    if (ctxOrFn instanceof Function) {
      t = mocha.it(title, ctxOrFn);
    } else {
      t = mocha.it(title, async function (this) {
        addContext(this, JSON.stringify(ctxOrFn));
        await (fn as mocha.AsyncFunc).call(this);
      });
    }
  } else {
    t = mocha.it(title);
  }
  return t;
}
