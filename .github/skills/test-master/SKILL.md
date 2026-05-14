---
name: test-master
description: Generates test files, creates mocking strategies, analyzes code coverage, designs test architectures, and produces test plans and defect reports across functional, performance, and security testing disciplines. Use when writing unit tests, integration tests, or E2E tests; creating test strategies or automation frameworks; analyzing coverage gaps; debugging flaky tests; or working on QA, regression, test automation, quality gates, or test maintenance. This project uses Mocha + Chai + Sinon (not Jest).
license: MIT
metadata:
  author: https://github.com/Jeffallan
  version: "1.1.1"
  domain: quality
  triggers: test, testing, QA, unit test, integration test, E2E, coverage, regression, test strategy, test automation, test framework, quality metrics, defect, flaky test, test maintenance
  role: specialist
  scope: testing
  output-format: report
  related-skills: debugging-wizard, code-reviewer
---

# Test Master

Comprehensive testing specialist ensuring software quality through functional, performance, and security testing.

> **Project note:** This repo uses **Mocha + Chai + Sinon** (not Jest/Vitest). Use `sinon.stub()` for mocking, `expect(...).to.equal(...)` for assertions, and `sinon.restore()` in `afterEach`. See `tests/unit/testHelper.ts` → `createMockContext()` for the standard context stub.

## Core Workflow

1. **Define scope** — Identify what to test and which testing types apply
2. **Create strategy** — Plan the test approach across unit, integration, and E2E tiers
3. **Write tests** — Implement tests with proper assertions
4. **Execute** — Run tests and collect results
   - If tests fail: classify the failure (assertion error vs. environment/flakiness), fix root cause, re-run
   - If tests are flaky: isolate ordering dependencies, check async handling, add retry or stabilization logic
5. **Report** — Document findings with severity ratings and actionable fix recommendations; verify coverage targets are met

## Test Structure (Mocha/Chai/Sinon)

```typescript
import { expect } from "chai";
import sinon from "sinon";
import { createMockContext } from "../../testHelper";
import { MyComponent } from "../../../src/myComponent";

describe("MyComponent", () => {
  describe("myMethod", () => {
    afterEach(() => {
      sinon.restore(); // always clean up stubs
    });

    it("should return ok when input is valid", async () => {
      const ctx = createMockContext();
      const result = await new MyComponent().myMethod(ctx, { key: "value" });
      expect(result.isOk()).to.be.true;
    });

    it("should return UserError when config is missing", async () => {
      const ctx = createMockContext();
      const result = await new MyComponent().myMethod(ctx, {});
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().name).to.equal("MissingConfigError");
    });
  });
});
```

## Constraints

**MUST DO**
- Test both `ok` and `err` paths for every `Result`-returning function
- Mock I/O boundaries (file system, HTTP, Azure SDK clients), not internal logic
- Use `sinon.stub()` / `sinon.sandbox`; always `sinon.restore()` in `afterEach`
- Name tests as `should <expected behavior> when <condition>`
- Run: `cd packages/<pkg> && npm run test:unit` (80% coverage gate)

**MUST NOT**
- Use Jest/Vitest matchers — this repo uses Chai's `expect(...).to` syntax
- Create order-dependent tests
- Test implementation details — test observable behaviour
- Ignore flaky tests — quarantine and fix them
- Use production data — use fixtures or `createMockContext()`

## Output Templates

When creating test plans, provide:
1. Test scope and approach
2. Test cases with expected outcomes
3. Coverage analysis
4. Findings with severity (Critical/High/Medium/Low)
5. Specific fix recommendations
