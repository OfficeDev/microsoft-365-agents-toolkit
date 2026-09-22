# ATK CLI Commands Reference

These commands apply to non-DA ATK projects and DAs routed to ATK by the [structural routing gate](declarative-agent-lifecycle.md#structural-routing-gate).

## Package and Validate

```bash
# Validate app
atk validate --env dev -i false

# Create app package
atk package --env dev -i false

# Sideload app
atk install --file-path ./appPackage.zip -i false

# Uninstall
atk uninstall --mode env --env dev --folder . -i false
```

## Share and Collaborate

```bash
# Share with entire tenant
atk share --scope tenant -i false

# Share with specific users
atk share --scope users --email 'user@example.com' -i false

# Grant collaborator access
atk collaborator grant -i false

# Check collaborator status
atk collaborator status
```

## Environment Management

```bash
# List environments
atk env list

# Add new environment
atk env add staging

# Reset environment
atk env reset --env dev -i false
```

## Troubleshooting

```bash
# Check system prerequisites
atk doctor

# Validate app manifest
atk validate --env dev -i false

# Upgrade project to latest toolkit version
atk upgrade -i false
```

**Port already in use:**

```powershell
# Windows: Find and kill process using port 3978
Get-NetTCPConnection -LocalPort 3978 | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

```bash
# macOS/Linux
lsof -ti:3978 | xargs kill -9
```

## Get Help

```bash
atk --help
atk new --help
```
