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
| `MaxDeactivationPercent` | no | Abort a full (`-Filter ALL`) `DiffSync`/`FullSync` run if more than this % of managed users would be deactivated/deleted in one go (default: `10`) |
| `PageSize` | no | UniFi API pagination (default: `200`) |
| `LogEnabled` | no | Write per-run log files (default: `true`) |
| `LogDir` | no | Log directory (default: `logs` next to script) |
| `AutoUpdate` | no | Check GitHub for a newer script on start (default: `false`) |
| `UpdateChannel` | no | Git branch to follow: `main` or `test` (default: `main`) |
| `UpdateRepoOwner` | no | GitHub user/org (default: `jeroen-vermeulen`) |
| `UpdateRepoName` | no | GitHub repo name (default: `salto-unifi-sync`) |
| `UpdateCheckIntervalHours` | no | Minimum hours between update checks (default: `24`) |

## Auto-update

When `AutoUpdate` is `true`, the script checks GitHub before each sync run:

1. Fetch `version.json` from the configured branch (`UpdateChannel`)
2. Compare remote version with the local `$ScriptVersion`
3. If newer: download script, verify SHA256, backup current file to `.bak`, replace, and **stop**
4. Print a message with the exact command to re-run (the current process still has the old version in memory)

Use **`main`** for production and **`test`** for pre-release builds.

```json
{
  "AutoUpdate": true,
  "UpdateChannel": "test",
  "UpdateRepoOwner": "jeroen-vermeulen",
  "UpdateRepoName": "salto-unifi-sync"
}
```

The repository is public. Downloads use the GitHub API without a token.

Emergency bypass:

```powershell
.\Sync-Salto-Unifi.ps1 -Mode DiffSync -Filter ALL -SkipUpdate
.\Sync-Salto-Unifi.ps1 -Mode DiffSync -Filter ALL -ForceUpdateCheck
```

If download or hash verification fails, the script logs `[UPDATE FAIL]` and continues with the current version.

## Safety check: mass deactivation/deletion

A wrong config value, a broken SQL connection string, or a SQL query that
unexpectedly returns zero rows can make every managed user look "ineligible"
at once. To avoid silently deactivating or deleting everyone in one run, a
full-population run (`-Filter ALL`) aborts *before* applying any change if
more than `MaxDeactivationPercent` (default `10`%) of the previously-managed
UniFi users would be deactivated or deleted.

- The impact (`X of Y managed user(s) (Z%)`) is always printed, including in
  `-Mode ShowDiff`, so you can review the plan first.
- Scoped runs (a single id, a range, or a list) are exempt — they are
  intentionally narrow (e.g. testing one user) and the percentage is not
  meaningful there.
- If you've reviewed the plan and the large change is genuinely correct
  (e.g. a bulk offboarding), re-run with `-Force` to override the check.

```powershell
.\Sync-Salto-Unifi.ps1 -Mode DiffSync -Filter ALL -Force
```

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

## What stays local

Never commit these files — they contain site-specific or sensitive data:

- `unifi-sync-config.json`
- `unifi-api.token` / any `*.token`
- `logs/`
- `Sync-Salto-Unifi.ps1.bak` (rollback backup after auto-update)

## Version

Current release: see [`version.json`](version.json) and `$ScriptVersion` inside the script.

## License

Public repository. Add a license file if you want a formal license.
