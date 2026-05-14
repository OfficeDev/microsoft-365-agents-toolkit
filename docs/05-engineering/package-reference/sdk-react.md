# `packages/sdk-react` — `@microsoft/teamsfx-react`

React hooks layer over `@microsoft/teamsfx` for Teams tab apps.

## Hooks

- `useTeams()` — Teams context (theme, locale, host info).
- `useTeamsFx()` — initialised TeamsFx instance.
- `useGraph()` — authenticated Graph client.
- `useData()` — generic async data hook with loading/error state.

## Targets

Browser only. Built with rollup. Peer-deps `react ^17 || ^18` and `@fluentui/react`.

## Co-location rationale

Same as [`sdk`](sdk.md) — stable contracts with templates, atomic updates.

## Documentation

End-user docs live at <https://aka.ms/teamsfx-docs>. This page is for contributors editing the React layer.
