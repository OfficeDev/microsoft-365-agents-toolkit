# Localization

Supported locales: **13 languages** in [`Localize/loc/`](../../Localize/loc/) — `cs`, `de`, `es`, `fr`, `it`, `ja`, `ko`, `pl`, `pt-BR`, `ru`, `tr`, `zh-Hans`, `zh-Hant`.

## Pipeline

| Step | Where |
|------|-------|
| English source | `package.nls.json` per package + per-package `src/resource/*.json` |
| Translation | `Localize/loc/{locale}/` (LocStudio output) |
| Build merge | Each package picks up `package.nls.{locale}.json` at build / runtime |
| Runtime lookup | `getLocalizedString(key, ...args)` (v3 `common/localizeUtils.ts`) · `Localizer.getString(key, ...args)` (v4 `core-next/src/localization/`) |

Both use Node's `util.format()` for parameter substitution. `getDefaultString(key, ...args)` returns the English source for logs and telemetry (always English).

## Authoring rules

- **Never concatenate** translated strings — fragments don't compose across languages.
- **Always parameterise**: `getLocalizedString("teamsfx.pluginX.failed", pluginName)` not `getLocalizedString("teamsfx.failed") + " " + pluginName`.
- Add new keys to `package.nls.json` first; translations follow in a separate PR after the LocStudio cycle.
- Keep keys descriptive: `teamsfx.{component}.{situation}` — never reuse keys across components.

## v4 differences

`core-next`'s `Localizer` is **injectable** (created via `createLocalizer()`), reusing the same `package.nls.json` format. This makes localised strings easy to mock in tests and supports per-tenant overrides if we ever need them. v4 retains backward compatibility with v3 string keys.

## Locale resolution

| Surface | Resolution |
|---------|-----------|
| VS Code | `vscode.env.language` |
| CLI | `LANG` / `LC_ALL` env vars; falls back to system default |
| Visual Studio | VS UI culture |

A locale not in the supported list falls back to English.
