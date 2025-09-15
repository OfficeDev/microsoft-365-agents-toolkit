Cherry-pick Summary for PR #14536
=====================================

This document tracks the completion of cherry-picking PR #14536 to the required branches.

## Original PR Details
- **PR Number**: #14536
- **Title**: "fix: reame link to teams ai v2"
- **Original Branch**: release/6.2
- **Commit SHA**: a7203cf25e9c0fee1642ad3f5366b8449df6ea58

## Cherry-pick Status
- [x] **dev branch**: Completed successfully
  - Local branch: `cherry-pick-to-dev`
  - Commit SHA: 985c71825313d451387545c94c147deda6e09f44d
  - Files changed: templates/vsc/js/command-and-response/README.md.tpl, templates/vsc/js/custom-copilot-rag-azure-ai-search/README.md.tpl, templates/vsc/ts/workflow/README.md.tpl

- [x] **release/VS1714P7 branch**: Completed successfully
  - Local branch: `cherry-pick-to-vs1714p7`  
  - Commit SHA: c70742bb37327daacf5b9fa2a2c74522e451474a
  - Files changed: templates/js/command-and-response/README.md.tpl, templates/js/custom-copilot-rag-azure-ai-search/README.md.tpl, templates/ts/workflow/README.md.tpl
  - Note: Required conflict resolution due to different path structures

## Changes Applied
All cherry-picks updated Teams AI SDK links from old documentation URLs to the standardized URL:
**New URL**: `https://aka.ms/teams-ai-library-v2`

## Verification
Both cherry-picks have been verified to apply the correct changes:
- Links properly updated to new Teams AI SDK V2 URL
- No additional unintended changes introduced
- All merge conflicts resolved appropriately for branch-specific path structures