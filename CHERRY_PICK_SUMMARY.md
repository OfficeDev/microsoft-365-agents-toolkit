# Cherry-pick Summary for PR #14321

## Overview
Successfully cherry-picked PR #14321 (`fix(SDK): Member.sendadaptivecard meet error`) to both target branches as requested in issue #14326.

## Changes Applied

### Dev Branch (Agents SDK version)
Applied commits:
- `0b16e7a0faf34d30f989804cd01a46b848881e05`: Main fix with dependency updates
- `bd1b08ea03ed3c946ce06a4f908bf89cf3eabd00`: Test coverage improvements

Files modified:
- `packages/sdk/CHANGELOG.md`: Version update 0.6.1 → 0.6.18
- `packages/sdk/package.json`: Agent SDK dependencies 0.6.1 → 0.6.18  
- `packages/sdk/pnpm-lock.yaml`: Lockfile updates
- `packages/sdk/src/conversationWithCloudAdapter/notification.ts`: ConnectorClient fallback + activity fix
- `packages/sdk/test/unit/node/conversationWithCloudAdapter/notification.spec.ts`: Test updates

### Release/VS1714P7 Branch (Bot Framework SDK version)
Applied equivalent fixes adapted for Bot Framework SDK:
- `packages/sdk/src/conversationWithCloudAdapter/notification.ts`: ConnectorClient fallback (different API)
- `packages/sdk/test/unit/node/conversationWithCloudAdapter/notification.spec.ts`: Test updates

## Key Differences Between Branches

### Dev Branch (Agents SDK)
- Uses Microsoft 365 Agents SDK (`@microsoft/agents-*` packages)
- API: `connectorClient.createConversationAsync(conversationParams)`
- Required `activity: undefined` fix for createConversation
- Has `paginate` method requiring additional parameter

### Release/VS1714P7 Branch (Bot Framework SDK)  
- Uses Bot Framework SDK (`botframework-*` packages)
- API: `connectorClient.conversations.createConversation(params)`
- No activity parameter needed
- Different pagination API (no additional parameter needed)

## Root Cause
Both branches had the same fundamental issue: `Member.sendAdaptiveCard()` method was failing when `connectorClient` was not available in the `TurnContext.turnState`. The fix adds a fallback to use `adapter.connectorClient` when the turnState version is unavailable.

## Testing
- Applied appropriate test updates to simulate the missing ConnectorClient scenario
- Tests verify the fallback behavior works correctly
- Changes are backward compatible

## Status
✅ Cherry-pick to dev branch: Complete
✅ Cherry-pick to release/VS1714P7 branch: Complete
✅ Tests updated for both branches: Complete
✅ No merge conflicts: Confirmed