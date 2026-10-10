# Operation - Update OAuth Configuration

- **Status:** Approved for implementation by the requester on 2026-10-09.
- **Action:** `oauth/update`.
- **Scope decision:** fx-core action contract only. This change does not add a Toolkit UI, CLI option, feature flag, or scaffolding behavior.

## Inputs

`supportedAccountTypes` has three update states:

| Input                  | Update meaning                                                    |
| ---------------------- | ----------------------------------------------------------------- |
| omitted                | Do not modify the service value and omit the property from PATCH. |
| `Enterprise`           | Explicitly disable personal Microsoft accounts.                   |
| `Enterprise, Consumer` | Explicitly enable personal Microsoft accounts.                    |

The parser and environment restrictions are the same as
[`oauth/register`](./register-oauth.md).

## Flow

```mermaid
flowchart TD
  Start[Execute oauth/update] --> Present{Parameter present?}
  Present -- no --> Compare[Compare existing OAuth fields only]
  Present -- yes --> Validate[Validate cloud, identity provider, and value]
  Validate --> CompareTypes[Compare normalized account types]
  Compare --> Changed{Any change?}
  CompareTypes --> Changed
  Changed -- no --> Skip[Skip PATCH]
  Changed -- yes --> Patch[PATCH changed registration]
```

## Boundary

- Does not infer or change account types when the parameter is omitted.
- Does not add UI, CLI, feature-flag, or scaffolding behavior.
- Does not support this field for Microsoft Entra SSO or sovereign clouds.

## Invariants

- Omission preserves the existing update behavior.
- A missing service value is semantically equivalent to `Enterprise`.
- PATCH carries the property only when the caller explicitly supplies it.

## Acceptance Criteria

| ID           | Runtime | Purpose               | Gate     | Harness                                      | Given / When                                                                              | Then                                                                      |
| ------------ | ------- | --------------------- | -------- | -------------------------------------------- | ----------------------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| OAUTH-MSA-05 | L1      | operation-integration | required | OAuth update driver with mocked TGS boundary | the parameter is omitted                                                                  | comparison and PATCH payload omit the property                            |
| OAUTH-MSA-06 | L1      | operation-integration | required | OAuth update driver with mocked TGS boundary | the caller explicitly changes the value                                                   | the change appears in confirmation and PATCH receives the canonical value |
| OAUTH-MSA-07 | L1      | compatibility         | required | OAuth update driver with mocked TGS boundary | the service omits the property and the caller supplies `Enterprise` with no other changes | the driver treats the values as equivalent and skips PATCH                |
| OAUTH-MSA-08 | L1      | operation-integration | required | OAuth update driver with mocked TGS boundary | invalid, Microsoft Entra, or unsupported-cloud input is supplied                          | the driver fails before PATCH                                             |
