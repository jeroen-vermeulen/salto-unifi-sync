# Salto → UniFi Access Sync

GitHub: [jeroen-vermeulen/salto-unifi-sync](https://github.com/jeroen-vermeulen/salto-unifi-sync)

One-way sync from **Salto ProAccess Space** (SQL Server) to **UniFi Access** (REST API).

Salto is the source of truth. The script creates/updates/deactivates/deletes **basic local** UniFi users that it manages (identified by `employee_number` = Salto `id_user` and `first_name` = `{id_user}_*`).

## Requirements

- Windows Server or PC with access to Salto SQL Server and UniFi Access API (LAN)
- PowerShell 5.1+
- `sqlcmd.exe` - either the classic ODBC-based tool (bundled with SQL Server / "Command Line Utilities for SQL Server") or the newer cross-platform `sqlcmd` (e.g. `winget install Microsoft.Sqlcmd`). The script auto-detects which one is on `PATH` and adjusts its arguments accordingly (see below); if both are installed, the highest-versioned one is used.
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
| `SqlServer` | yes | SQL Server to read the Salto database from: `localhost\SQLEXPRESS`, `host\instance`, `host,port`, or a LocalDB instance `(localdb)\<instance>` (Salto ProAccess Space's `service.ini` lists it as `DBServerName`) |
| `Database` | yes | Salto database name (e.g. `SALTO_SPACE`) |
| `UserGroupName` | yes | UniFi user group for access |
| `SaltoUserType` | no | Filter on user type (default: `STAFF`) |
| `OnlyActive` | no | Only Salto users with status ACTIVE (default: `true`) |
| `RequireTag` | no | Skip users without active NFC tag (default: `true`) |
| `DeactivateWhenIneligible` | no | Cleanup script-managed users in UniFi (default: `true`) |
| `DeleteOrphanNfcTokens` | no | Remove NFC tokens from inventory on user delete (default: `true`) |
| `MaxDeactivationPercent` | no | Abort a full (`-Filter ALL`) `DiffSync`/`FullSync` run if more than this % of managed users would be deactivated/deleted in one go (default: `10`) |
| `SqlPreflight` | no | Check up front that `SqlServer` can answer (see "SQL Server preflight"); set `false` to skip (default: `true`) |
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

## sqlcmd variant detection

Two incompatible tools are both called `sqlcmd.exe` in the wild:

- **Classic** (ODBC-based, bundled with SQL Server / "Command Line Utilities for SQL Server"). Supports `-f i:<codepage>,o:<codepage>` for forcing UTF-8 I/O.
- **Modern** cross-platform `sqlcmd` (e.g. installed via `winget install Microsoft.Sqlcmd`). Does not support `-f` and errors with `Sqlcmd: 'f': Unknown Option` if it's passed — it is UTF-8 by default so the flag isn't needed there.

On startup the script runs `Get-Command sqlcmd -All` to find every `sqlcmd` on `PATH`, checks each one's help output for `-f <codepage>` to classify it, and only adds `-f i:65001,o:65001` when the resolved binary supports it. If more than one `sqlcmd` is found (e.g. both variants installed side by side), the **highest-versioned** one is selected. The chosen path and detected variant are printed at the start of the run:

```
Using sqlcmd: C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\SQLCMD.EXE [classic (ODBC-based)]
```

## SQL Server preflight

A wrong `SqlServer` (typically a config copied from another machine) used to surface as a
slow, cryptic `sqlcmd` timeout such as `Timed out waiting for pipe '\\.\pipe\SQLLocal\...'`.
Since v1.3.16 the script checks first and stops within a second with an actionable message:

| `SqlServer` form | What is checked |
|------------------|-----------------|
| `localhost\NAME`, `.\NAME`, `(local)`, `<this computer>` | The Windows service `MSSQL$NAME` (`MSSQLSERVER` for the default instance) exists and is running. The message lists the SQL services that do exist. |
| `(localdb)\NAME` | `SqlLocalDB.exe i` lists `NAME` for the account running the script. LocalDB instances belong to a single Windows account, so a service-owned instance is invisible to another user: run the script as the owning account (normally the Salto service account). Shared instances (`(localdb)\.\name`) are not checked. |
| `host,port` (or `localhost,port`) | The TCP port answers (3 s timeout). |
| `host` or `host\NAME` (remote) | The host name resolves; for a default instance TCP 1433 must answer. Named-instance ports come from SQL Browser (UDP), so only the name is checked. |

If the preflight ever misjudges a working setup, set `"SqlPreflight": false`. It does not
test logins or permissions; those errors still come from `sqlcmd` itself.

## NFC tag replacement detection

Each script-managed UniFi user gets an NFC token aliased `salto-{id_user}` in the UniFi
token inventory. Up to v1.3.14, two places treated *owning that alias* as proof the
user's current card was correct:

- `Test-NfcInSync` returned "in sync" if any of the user's attached cards had a token
  aliased `salto-{id_user}` — without checking that token's actual UID matched the
  Salto-desired tag.
- `Ensure-NfcToken` matched an *existing* token on alias alone, so even after the above
  was fixed it would still hand back the old card's token instead of importing one for
  the new tag.

Together these meant: once a user had a script-managed card, replacing their physical
Salto pass was silently ignored forever — `DiffSync` never proposed `UPDATE_NFC` for
them, since the alias (not the tag) was what the old pass check looked at, and the alias
never changes.

Since v1.3.15: `Test-NfcInSync` only trusts the alias if the UID currently behind it
matches the desired tag, and `Ensure-NfcToken` matches *only* on tag value when looking
for an existing token. When a user's alias is still attached to a superseded token (old
tag no longer current), that token is unassigned from the user and deleted from the UniFi
token inventory, then a fresh token is imported for the new tag under the same alias — so
`salto-{id_user}` always resolves to exactly one, current token.

## Performance

Every run indexes the UniFi user group once against the UniFi Access API. Up to v1.3.13
this involved two per-user redundancies:

- `Build-UniFiUserIndex` fetched `GET /users/{id}` individually for every UniFi user,
  even though the bulk `GET /users` listing (already fetched for pagination) returns the
  exact same fields (`employee_number`, `first_name`, `last_name`, `status`, `nfc_cards`,
  `username`, `email`, `user_email`).
- `Test-UserInGroup` re-fetched the *entire* group's member list from scratch for every
  Salto user matched to an existing UniFi account, instead of once for the whole run.

Since v1.3.14 both are eliminated: the UniFi index is built directly from the bulk list,
and group membership is fetched once and checked via an in-memory lookup. This removes
one `curl.exe` process + TLS handshake per UniFi user, plus one per matched user for the
group check — with no change in output (verified with a byte-for-byte diff of the plan
table before/after on a live instance). Expect `-Mode ShowDiff -Filter ALL` to run
noticeably faster, especially with a larger UniFi user base.

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
