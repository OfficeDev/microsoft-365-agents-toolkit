# Release Schedule

This document defines the release schedule for Microsoft 365 Agents Toolkit products.

## CY26 Q1

| Products | Release Type | Version | Cut Bits Date | Status | Branch | preid | series |
|----------|--------------|---------|---------------|--------|--------|-------|--------|
| VSC  | Prerelease   | 6.7.2026012001 | 2026-01-22 |  | release/6.7 | preview | CY260120 |
| VSC  | Prerelease   | 6.7.2026012002 | 2026-01-22 |  | release/6.7 | preview | CY260121 |

## Automation Configuration

To enable automated release branch creation and deployment, define the CD parameters directly in the schedule.
These values are authoritative (the automation does not infer them).

For rows that should be automated, include these columns in the table:
- `Branch`
- `preid`
- `series`

Products column has to be either `VSC` or `VS`.

The automation workflow will:
- Parse this schedule
- Create release branches based on the `Branch` column
- Trigger CD pipeline with the provided parameters (`preid`, `series`) and derived `vsrelease`
- Require manual approval before execution

Notes:
- `Cut Bits Date` must be in ISO format `YYYY-MM-DD` (used for automatic selection).
- `vsrelease` is derived from `Products`: `VS` => true, `VSC` => false.
