Test your agent in the actual Teams environment. Requires M365 account and HTTPS endpoint.

# Prerequisite

User has Microsoft 365 account with sideloading enabled:
Use `atk account show` to check and `atk account login m365` to login.
Use `atk doctor` to check sideloading permission.

# Opening in Different Hosts

Get your app IDs from `env/.env.local`, then open:

| Host | URL |
|------|-----|
| Teams web | `https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TENANT_ID}}&login_hint=${{USER_EMAIL}}` |
| Outlook web | `https://outlook.office.com/host/${{M365_APP_ID}}` |
| Office web | `https://www.office.com/m365apps/${{M365_APP_ID}}` |

# Dev Tunnels for Bots

**IMPORTANT**: For bot projects, you must start a public devtunnel BEFORE provisioning.

The tunnel must be public/anonymous so Teams can reach your bot:
```bash
# Start a public devtunnel
devtunnel host -p 3978 --allow-anonymous
```

Then set `BOT_ENDPOINT` in `env/.env.local` with the tunnel URL before running `atk provision`.

# Troubleshooting

## Bot not responding

1. Verify bot is running on specified endpoint
2. Check bot logs for errors
3. Ensure your bot endpoint is accessible:
   ```bash
   curl http://localhost:3978/api/messages
   ```

## Teams shows "app not available"

This usually means BOT_ENDPOINT requires HTTPS. Use Agents Playground instead, or ensure dev tunnel is running and BOT_ENDPOINT is properly configured.
