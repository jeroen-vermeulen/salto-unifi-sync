# Salto → UniFi Access Sync

GitHub: [jeroen-vermeulen/salto-unifi-sync](https://github.com/jeroen-vermeulen/salto-unifi-sync)

One-way sync from **Salto ProAccess Space** (SQL Server) to **UniFi Access** (REST API).

Salto is the source of truth. The script creates/updates/deactivates/deletes **basic local** UniFi users that it manages (identified by `employee_number` = Salto `id_user` and `first_name` = `{id_user}_*`).

## Requirements

- Windows Server or PC with access to Salto SQL Server and UniFi Access API (LAN)
- PowerShell 5.1+
- `sqlcmd.exe` (SQL Server tools)
- `curl.exe` (included with Windows 10/11 and Server 2019+)

## Quick start

1. Clone or copy this repo to your server, e.g. `C:\Tools\Sync-Salto-Unifi\`
   ```powershell
   git clone https://github.com/jeroen-vermeulen/salto-unifi-sync.git C:\Tools\Sync-Salto-Unifi
   ```
2. Copy the config template:
   ```powershell
   Copy-Item unifi-sync-config.json.example unifi-sync-config.json
   ```
3. Edit `unifi-sync-config.json` — set `UnifiHost`, `ApiTokenFile`, `SqlServer`, `UserGroupName`, etc.
4. Create your UniFi API token file (one line, no BOM):
   ```powershell
   Set-Content -Path C:\path\to\unifi-api.token -Value 'YOUR_TOKEN' -NoNewline
   ```
5. Dry run:
   ```powershell
   .\Sync-Salto-Unifi.ps1 -Mode ShowDiff -Filter ALL
   ```
6. Apply changes:
   ```powershell
   .\Sync-Salto-Unifi.ps1 -Mode DiffSync -Filter ALL
   ```

## Configuration

| Key | Required | Description |
|-----|----------|-------------|
| `UnifiHost` | yes | UniFi Access API base URL |
| `ApiTokenFile` | yes | Path to bearer token file (local, not in git) |
| `SqlServer` | yes | SQL Server instance |
| `Database` | yes | Salto database name (e.g. `SALTO_SPACE`) |
| `UserGroupName` | yes | UniFi user group for access |
| `SaltoUserType` | no | Filter on user type (default: `STAFF`) |
| `OnlyActive` | no | Only Salto users with status ACTIVE (default: `true`) |
| `RequireTag` | no | Skip users without active NFC tag (default: `true`) |
| `DeactivateWhenIneligible` | no | Cleanup script-managed users in UniFi (default: `true`) |
| `DeleteOrphanNfcTokens` | no | Remove NFC tokens from inventory on user delete (default: `true`) |
| `PageSize` | no | UniFi API pagination (default: `200`) |
| `LogEnabled` | no | Write per-run log files (default: `true`) |
| `LogDir` | no | Log directory (default: `logs` next to script) |

## Modes

| Mode | Description |
|------|-------------|
| `ShowDiff` | Show planned actions, no changes |
| `DiffSync` | Apply only required changes (default) |
| `FullSync` | Re-apply all mapped fields for filtered users |

## Filter

| Value | Meaning |
|-------|---------|
| `ALL` | All eligible Salto users |
| `1000` | Single `id_user` |
| `10-20` | Range (inclusive) |
| `"100,200,300"` | Comma-separated list (quote in PowerShell) |

## Branches

| Branch | Raw `version.json` URL |
|--------|------------------------|
| `main` (production) | `https://raw.githubusercontent.com/jeroen-vermeulen/salto-unifi-sync/main/version.json` |
| `test` (pre-release) | `https://raw.githubusercontent.com/jeroen-vermeulen/salto-unifi-sync/test/version.json` |

> Auto-update from GitHub (channel selection via config) will be added in a future release.

## What stays local

Never commit these files — they contain site-specific or sensitive data:

- `unifi-sync-config.json`
- `unifi-api.token` / any `*.token`
- `logs/`

## Version

Current release: see [`version.json`](version.json) and `$ScriptVersion` inside the script.

## License

Private / internal use — add a license if you publish publicly.
