# 0.1.0

Initial release of the SDK for React Hook.

- Add 3 React Hook functions: useTeamsFx, useGraph and useData

# 2.0.0

2.0.0 is a major version upgrade for TeamsFx SDK for React hooks with many improvements and supports.

## Added

- Add `useTeams` hook, referenced [msteams-react-base-component](https://github.com/wictorwilen/msteams-react-base-component).
- Update to use TeamsFx SDK 2.0, which supports [Teams JS SDK 2.0](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-sdk?tabs=javascript%2Cmanifest-teams-toolkit).
- Support both CJS and ESM modules now, so that webpack and other tools can do proper tree-shaking and optimizations.

# 3.0.0

3.0.0 is a major version upgrade for TeamsFx SDK for React hooks migrating to support React 18 and Fluent UI v9.

## Added

- Update to support [React 18](https://reactjs.org/).
- Migrate to use [Fluent UI v9](https://react.fluentui.dev/?path=/docs/concepts-introduction--page), which includes breaking changes of `useTeams`, `useTeamsUserCredential` and `useTeamsFx`.

# 3.1.0-alpha

- Add `loading` parameter to the return value of `useTeams` hook.

# 3.1.2
- Update peer dependency `@microsoft/teams-js` version to `^2.19.0`.


# 4.0.3
- Deprecation notice: This package will be deprecated by 2026-07. Users can follow the migration guide to migrate their apps.
