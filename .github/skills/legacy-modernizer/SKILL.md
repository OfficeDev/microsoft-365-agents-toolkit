---
name: legacy-modernizer
description: Designs incremental migration strategies, identifies service boundaries, produces dependency maps and migration roadmaps, and generates API facade designs for aging codebases. Use when modernizing legacy systems, implementing strangler fig pattern or branch by abstraction, decomposing monoliths, upgrading frameworks or languages, or reducing technical debt without disrupting business operations.
license: MIT
metadata:
  author: https://github.com/Jeffallan
  version: "1.1.0"
  domain: specialized
  triggers: legacy modernization, strangler fig, incremental migration, technical debt, legacy refactoring, system migration, legacy system, modernize codebase, v3 to v4, migration
  role: specialist
  scope: architecture
  output-format: code+analysis
  related-skills: test-master, devops-engineer
---

# Legacy Modernizer

## Core Workflow

1. **Assess system** — Analyze codebase, dependencies, risks, and business constraints. Produce a dependency map and risk register before proceeding.
   - *Validation checkpoint:* Confirm all external integrations and data contracts are documented before moving to step 2.

2. **Plan migration** — Design an incremental roadmap with explicit rollback strategies per phase.
   - *Validation checkpoint:* Confirm each phase has a defined rollback trigger and owner.

3. **Build safety net** — Create characterization tests and monitoring before touching production code. Target 80%+ coverage of existing behavior.
   - *Validation checkpoint:* Run the characterization test suite and confirm it passes green on the unmodified legacy system before proceeding.

4. **Migrate incrementally** — Apply strangler fig pattern with feature flags. Route traffic via a facade; shift load gradually.
   - *Validation checkpoint:* Verify error rates and metrics remain within baseline thresholds after each increment.

5. **Validate & iterate** — Run full test suite, review monitoring, and confirm business behavior is preserved before retiring legacy code.
   - *Validation checkpoint:* New code must be proven stable at 100% traffic for at least one release cycle before legacy path is removed.

## Code Examples

### Strangler Fig Facade (TypeScript)
```typescript
// facade.ts — routes to legacy or new implementation based on a feature flag
import { featureFlagManager, FeatureFlagName } from "../common/featureFlags";
import { LegacyOrderService } from "./legacy/orderService";
import { NewOrderService } from "./next/orderService";

export class OrderServiceFacade {
  private legacy = new LegacyOrderService();
  private next = new NewOrderService();

  getOrder(orderId: string) {
    if (featureFlagManager.getBooleanValue(FeatureFlagName.V4OrderService, false)) {
      return this.next.fetch(orderId);
    }
    return this.legacy.get(orderId);
  }
}
```

### Characterization Test Template (Mocha/Chai)
```typescript
// Captures existing legacy behavior as a golden-master safety net.
import { expect } from "chai";
import { LegacyOrderService } from "../../src/legacy/orderService";

describe("LegacyOrderService — characterization", () => {
  const service = new LegacyOrderService();

  const cases = [
    { orderId: "ORD-001", expectedStatus: "SHIPPED" },
    { orderId: "ORD-002", expectedStatus: "PENDING" },
  ];

  cases.forEach(({ orderId, expectedStatus }) => {
    it(`preserves status=${expectedStatus} for ${orderId}`, async () => {
      const result = await service.get(orderId);
      expect(result.status).to.equal(expectedStatus);
    });
  });
});
```

## Constraints

### MUST DO
- Maintain zero production disruption during all migrations
- Create comprehensive test coverage before refactoring (target 80%+)
- Use feature flags for all incremental rollouts
- Implement monitoring and rollback procedures
- Document all migration decisions and rationale
- Preserve existing business logic and behavior

### MUST NOT DO
- Big bang rewrites or replacements
- Skip testing legacy behavior before changes
- Deploy without rollback capability
- Break existing integrations or APIs
- Remove legacy code before new code is proven

## Output Templates

When implementing modernization, provide:
1. Assessment summary (risks, dependencies, approach)
2. Migration plan (phases, rollback strategy, metrics)
3. Implementation code (facades, adapters, new services)
4. Test coverage (characterization, integration, e2e)

## Knowledge Reference

Strangler fig pattern, branch by abstraction, characterization testing, incremental migration, feature flags, canary deployments, API versioning, zero-downtime deployment
