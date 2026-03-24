# dev.debug-test-ts

## purpose

Developer tools, local debugging, DevTools plugin, sideloading, tunneling, ConsoleLogger configuration, and build verification for Teams SDK v2.

## rules

1. Always include `DevtoolsPlugin` from `@microsoft/teams.dev` in the `plugins` array during local development. It provides a web-based DevTools UI, WebSocket-based real-time activity inspection, and message replay capabilities. [github.com/microsoft/teams.ts -- dev](https://github.com/microsoft/teams.ts/tree/main/packages/dev)
2. The DevTools UI runs at `http://localhost:{PORT+1}/devtools` (default `http://localhost:3979/devtools`). It starts automatically when the app starts with `DevtoolsPlugin` registered. Open it in a browser to inspect inbound and outbound activities. [github.com/microsoft/teams.ts -- dev](https://github.com/microsoft/teams.ts/tree/main/packages/dev)
3. The bot endpoint is `http://localhost:{PORT}/api/messages` (default `http://localhost:3978/api/messages`). This is the URL that Teams (or the Bot Framework Emulator) sends activities to. For Azure deployment, update the messaging endpoint to `https://your-domain/api/messages`. [github.com/microsoft/teams.ts](https://github.com/microsoft/teams.ts)
4. Run locally with `npm run dev` which executes `tsx watch -r dotenv/config src/index.ts`. This provides TypeScript execution with automatic file watching -- changes to source files trigger an instant restart without a build step. [github.com/microsoft/teams.ts](https://github.com/microsoft/teams.ts)
5. Configure logging with `ConsoleLogger` at the appropriate level: `'trace'` for deep SDK internals, `'debug'` for development, `'info'` for staging, `'warn'` or `'error'` for production. Use `pattern: '-azure/msal-node'` to suppress noisy MSAL authentication logs. Child loggers are created with `logger.child('name')`. [github.com/microsoft/teams.ts -- common](https://github.com/microsoft/teams.ts/tree/main/packages/common)
6. Run `npx tsc --noEmit` as a build verification gate before testing or deploying. This type-checks all TypeScript source without producing output files. The project must compile cleanly -- type errors caught here prevent runtime failures. [github.com/microsoft/teams.ts](https://github.com/microsoft/teams.ts)
7. For testing with real Teams clients locally, use a tunneling solution (ngrok, dev tunnels, Cloudflare Tunnel) to expose your local `localhost:3978` endpoint over HTTPS. Update the Azure Bot messaging endpoint to the tunnel URL (e.g., `https://abc123.ngrok.io/api/messages`). [learn.microsoft.com -- Dev tunnels](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/overview)
8. Sideload the app by creating a zip of the `appPackage/` directory (manifest.json + icons) and uploading in Teams via Apps > Manage your apps > Upload a custom app. Sideloading requires admin permission or a developer tenant. [learn.microsoft.com -- Sideload apps](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload)
9. Use `skipAuth: true` in `AppOptions` for purely local development against DevTools without Azure Bot credentials. This disables JWT validation. Never use it in production or when testing against real Teams clients. [github.com/microsoft/teams.ts -- apps](https://github.com/microsoft/teams.ts/tree/main/packages/apps)
10. The M365 Agents Toolkit VS Code extension provides GUI-based provisioning, dev tunnel management, one-click deployment, and manifest editing. It is an alternative to manual CLI workflows for developers who prefer a visual interface. [learn.microsoft.com -- Agents Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/toolkit-v4/teams-toolkit-fundamentals-vs)

## patterns

### Local development setup with DevTools

```typescript
import { App } from '@microsoft/teams.apps';
import { ConsoleLogger } from '@microsoft/teams.common';
import { DevtoolsPlugin } from '@microsoft/teams.dev';

const app = new App({
  // Use skipAuth for local DevTools-only testing (no Azure credentials needed)
  // Remove skipAuth when testing against real Teams clients
  skipAuth: true,

  logger: new ConsoleLogger('dev-bot', { level: 'debug' }),
  plugins: [new DevtoolsPlugin()],
});

app.on('message', async ({ reply, activity }) => {
  await reply(`Echo: ${activity.text}`);
});

// Bot endpoint: http://localhost:3978/api/messages
// DevTools UI:  http://localhost:3979/devtools
app.start(3978);
```

### Production-ready logging configuration

```typescript
import { App } from '@microsoft/teams.apps';
import { ConsoleLogger } from '@microsoft/teams.common';
import { DevtoolsPlugin } from '@microsoft/teams.dev';

// Development logger: verbose, noisy auth logs suppressed
const devLogger = new ConsoleLogger('my-bot', {
  level: 'debug',
  pattern: '-azure/msal-node',
});

// Production logger: only warnings and errors
const prodLogger = new ConsoleLogger('my-bot', {
  level: 'warn',
});

// Choose based on environment
const isProduction = process.env.NODE_ENV === 'production';
const logger = isProduction ? prodLogger : devLogger;

const app = new App({
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  tenantId: process.env.TENANT_ID,
  logger,
  // Only include DevtoolsPlugin in non-production
  plugins: isProduction ? [] : [new DevtoolsPlugin()],
});

// Child loggers for scoped output
app.on('message', async (ctx) => {
  const handlerLog = ctx.log; // Scoped to this activity
  handlerLog.info(`Message from ${ctx.activity.from.name}`);
  // Output: [my-bot] Message from John Doe

  await ctx.send('Hello!');
});

app.start(process.env.PORT || 3978).catch(console.error);
```

### Build verification and debugging workflow

```typescript
// Step-by-step local development workflow:

// 1. Install dependencies
//    npm install

// 2. Create .env with required variables
//    CLIENT_ID=<your-bot-client-id>
//    CLIENT_SECRET=<your-bot-client-secret>
//    TENANT_ID=<your-azure-tenant-id>
//    PORT=3978

// 3. Type-check the project (build gate)
//    npx tsc --noEmit
//    Fix any type errors before proceeding.

// 4. Start dev server with file watching
//    npm run dev
//    This runs: tsx watch -r dotenv/config src/index.ts

// 5. Open DevTools in browser
//    http://localhost:3979/devtools
//    Send test messages, inspect activities in real time.

// 6. For real Teams testing, set up a tunnel:
//    ngrok http 3978
//    Copy the HTTPS URL (e.g., https://abc123.ngrok.io)
//    Update Azure Bot messaging endpoint to:
//      https://abc123.ngrok.io/api/messages

// 7. Sideload the app:
//    cd appPackage
//    zip -r ../mybot.zip manifest.json color.png outline.png
//    Upload mybot.zip in Teams > Apps > Upload a custom app

// 8. Build for production:
//    npm run build     # Compiles to dist/ via tsup
//    npm run start     # Runs compiled JS: node -r dotenv/config .

// Common troubleshooting:
// - Bot not responding in Teams?
//     Check: tunnel running, messaging endpoint updated, manifest scopes correct
// - DevTools blank?
//     Check: DevtoolsPlugin in plugins array, port+1 not blocked
// - Type errors on import?
//     Check: tsconfig module is "NodeNext", not "commonjs"
// - .env not loaded?
//     Check: dotenv in devDependencies, -r dotenv/config in scripts
```

## pitfalls

- **Forgetting to start the tunnel**: Without ngrok or dev tunnels, the Azure Bot Framework cannot reach your local endpoint. Teams messages never arrive. Always verify the tunnel is running and the messaging endpoint is updated.
- **DevTools port conflict**: DevTools runs on `PORT + 1`. If port 3979 is already in use, DevTools fails silently. Check for port conflicts or change the bot's `PORT`.
- **Using `skipAuth` with real Teams clients**: `skipAuth: true` disables JWT validation. Real Teams activities require proper authentication. Use `skipAuth` only with DevTools for rapid iteration.
- **Not running `npx tsc --noEmit`**: Skipping the type-check means errors surface only at runtime. Always run this gate after changes, especially before committing or deploying.
- **Stale tunnel URL in Azure Bot config**: Ngrok generates a new URL each time it restarts (unless on a paid plan). Forgetting to update the Azure Bot messaging endpoint after restarting ngrok means the bot stops receiving messages.
- **Missing sideload permissions**: Sideloading requires either admin-enabled custom app upload or a Microsoft 365 developer tenant. Without it, the "Upload a custom app" option does not appear in Teams.
- **Wrong log level in production**: Running with `'debug'` or `'trace'` in production floods logs and can impact performance. Switch to `'info'` or `'warn'` for deployed environments.
- **Testing only in DevTools**: DevTools simulates a Teams client but does not replicate all Teams behaviors (e.g., @mention stripping, SSO token exchange, card rendering differences). Always test in a real Teams client before shipping.

## references

- [Teams SDK v2 -- @microsoft/teams.dev (DevtoolsPlugin)](https://github.com/microsoft/teams.ts/tree/main/packages/dev)
- [Teams SDK v2 GitHub repository](https://github.com/microsoft/teams.ts)
- [Teams: Sideload apps](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload)
- [Azure Dev Tunnels](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/overview)
- [ngrok documentation](https://ngrok.com/docs)
- [M365 Agents Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/toolkit-v4/teams-toolkit-fundamentals-vs)
- [Teams: Test and debug](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/debug)

## instructions

This expert covers local development, debugging, and testing workflows for Teams SDK v2 bots. Use it when you need to:

- Set up `DevtoolsPlugin` and access the DevTools UI at `localhost:3979/devtools`
- Run the bot locally with `npm run dev` (tsx watch with hot reload)
- Configure `ConsoleLogger` levels and noise filtering for different environments
- Use `skipAuth: true` for credential-free local testing
- Set up ngrok or dev tunnels for testing with real Teams clients
- Sideload the bot by packaging and uploading `appPackage/` as a zip
- Run `npx tsc --noEmit` as a build verification gate
- Troubleshoot common issues (bot not responding, DevTools blank, import errors)
- Understand the difference between DevTools testing and real Teams testing

Pair with `runtime.app-init-ts.md` for App constructor setup and `project.scaffold-files-ts.md` for npm scripts and project file structure. Pair with `project.scaffold-files-ts.md` for npm scripts and build verification, and `runtime.app-init-ts.md` for DevtoolsPlugin configuration.

## research

Deep Research prompt:

"Write a micro expert on developing, debugging, and testing Teams SDK v2 bots in TypeScript. Cover DevtoolsPlugin setup from @microsoft/teams.dev, DevTools UI at localhost:3979/devtools with WebSocket activity inspection and message replay, local development with npm run dev (tsx watch), bot endpoint at localhost:3978/api/messages, ConsoleLogger configuration (levels: error/warn/info/debug/trace, pattern filtering, child loggers), skipAuth for local testing, ngrok and dev tunnels for HTTPS exposure, sideloading via zip upload, npx tsc --noEmit build gate, M365 Agents Toolkit VS Code extension, and a troubleshooting flowchart for common issues (bot not responding, port conflicts, stale tunnel URLs, missing sideload permissions). Include a step-by-step local workflow and production logging patterns."
