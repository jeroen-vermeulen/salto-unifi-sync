<#
.SYNOPSIS
    One-way sync from Salto ProAccess Space (SALTO_SPACE) to UniFi Access.

.DESCRIPTION
    Salto is the source of truth. Only basic local UniFi users managed by this
    script (matched via employee_number = Salto id_user) are created or updated.
    Non-basic UniFi accounts (e.g. with email) are not updated in place. If they have
    no NFC tag or the wrong tag, a new script-managed basic user is created instead.
    If a non-basic account already has the correct Salto tag, it is left unchanged.
    Script-managed Basic users (employee_number + {id_user}_ first_name prefix) are
    deactivated in UniFi when no longer eligible in Salto (removed, inactive, or no tag).

    Mapping:
      UniFi first_name  = "{id_user}_{Salto FirstName}"
      UniFi last_name   = Salto LastName
      UniFi employee_number = Salto id_user (3rd-party / external reference)
      NFC card          = Salto ROMCode (TagID)
      User group        = configured in UserGroupName

.PARAMETER Mode
    DiffSync  - Apply only required changes (default)
    FullSync  - Re-apply all mapped fields for every filtered Salto user
    ShowDiff  - Show planned actions without making changes

.PARAMETER Filter
    ALL       - All eligible Salto users (default)
    10-20     - id_user between 10 and 20 (inclusive)
    1000      - Single id_user
    100,200,300 - Comma-separated id_user list (quote in PowerShell: "100,200,300")

.PARAMETER ConfigPath
    Path to JSON config. Default: unifi-sync-config.json next to this script.

.PARAMETER SkipUpdate
    Skip the GitHub auto-update check.

.PARAMETER ForceUpdateCheck
    Check GitHub for updates even if UpdateCheckIntervalHours has not elapsed.

.PARAMETER Force
    Apply the sync even if the deactivate/delete safety check (MaxDeactivationPercent,
    default 10%) would otherwise abort the run. Only use after reviewing the plan
    with -Mode ShowDiff.

.PARAMETER ShowAll
    Also list the users that need no change in the plan. By default the plan lists only
    the users that would be changed, with what changes (old -> new) per user.

.PARAMETER NoTranscript
    Internal. Used by the delegated run started through "SqlRunAs": the starting
    (elevated) session writes the log file, so this process does not start its own.

.PARAMETER ParentLogPath
    Internal. Log file path written by the starting session, shown in the run header.

.PARAMETER NoDelegate
    Run in the current account even when the config sets "SqlRunAs". Normally the
    script starts itself as the SqlRunAs account (needed for a private LocalDB
    instance owned by another account) and uses this switch for that inner run.

.EXAMPLE
    .\Sync-Salto-Unifi.ps1 -Mode ShowDiff -Filter ALL

.EXAMPLE
    .\Sync-Salto-Unifi.ps1 -Mode DiffSync -Filter 1000

.EXAMPLE
    .\Sync-Salto-Unifi.ps1 -Mode FullSync -Filter 10-20
#>
[CmdletBinding()]
param(
    [ValidateSet('DiffSync', 'FullSync', 'ShowDiff')]
    [string]$Mode = 'DiffSync',

    [string]$Filter = 'ALL',

    [string]$ConfigPath = '',

    [switch]$SkipUpdate,

    [switch]$ForceUpdateCheck,

    [switch]$Force,

    [switch]$ShowAll,

    [switch]$NoDelegate,

    [switch]$NoTranscript,

    [string]$ParentLogPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $PSScriptRoot) {
    $PSScriptRoot = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
}
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot 'unifi-sync-config.json'
}
$ScriptVersion = '1.3.23'

$script:RunLogPath = $null
$script:TranscriptActive = $false
$script:LogProblem = $null

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-JsonNoBom([string]$Path, [string]$Content) {
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8)
}

function Get-ObjProp($Obj, [string]$Name) {
    if ($null -eq $Obj) { return $null }
    $prop = $Obj.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function Set-ConfigDefault {
    param(
        [Parameter(Mandatory)]$Cfg,
        [Parameter(Mandatory)][string]$Name,
        $Value
    )
    if ($null -eq (Get-ObjProp $Cfg $Name)) {
        $Cfg | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}

function Get-ConfigBool {
    param(
        $Cfg,
        [Parameter(Mandatory)][string]$Name,
        [bool]$Default = $false
    )
    $val = Get-ObjProp $Cfg $Name
    if ($null -eq $val) { return $Default }
    return [bool]$val
}

function Normalize-UnicodeText([string]$Value) {
    if ($null -eq $Value) { return '' }
    $trimmed = $Value.Trim()
    if ($trimmed.Length -eq 0) { return '' }
    return $trimmed.Normalize([System.Text.NormalizationForm]::FormC)
}

function Normalize-SqlCmdValue([string]$Value) {
    if ($null -eq $Value) { return '' }
    $trimmed = $Value.Trim()
    if ($trimmed.Length -eq 0) { return '' }
    if ($trimmed -match '^(?i)NULL$') { return '' }
    return $trimmed
}

function Test-NameDiffers {
    param(
        [string]$UniFirst,
        [string]$UniLast,
        [string]$DesiredFirst,
        [string]$DesiredLast
    )
    $cmp = [System.StringComparison]::Ordinal
    $firstDiff = -not [string]::Equals(
        (Normalize-UnicodeText $UniFirst),
        (Normalize-UnicodeText $DesiredFirst),
        $cmp
    )
    $lastDiff = -not [string]::Equals(
        (Normalize-UnicodeText $UniLast),
        (Normalize-UnicodeText $DesiredLast),
        $cmp
    )
    return ($firstDiff -or $lastDiff)
}

$script:SqlCmdResolved = $false
$script:SqlCmdPath = $null
$script:SqlCmdSupportsCodePage = $false
$script:LocalDbPipeCache = @{}

function Get-SqlCmdFileVersion([string]$Path) {
    try {
        $v = (Get-Item -LiteralPath $Path).VersionInfo.FileVersionRaw
        if ($v) { return $v }
    } catch { }
    return [version]'0.0.0.0'
}

function Test-SqlCmdSupportsCodePage([string]$Path) {
    # Classic ODBC-based sqlcmd (Command Line Utilities for SQL Server) supports
    # -f i:<codepage>,o:<codepage>. The newer cross-platform "go-sqlcmd" (e.g. from
    # `winget install Microsoft.Sqlcmd`) does not know this flag and errors with
    # "Unknown Option" - it is UTF-8 safe by default so the flag is not needed there.
    try {
        $help = (& $Path '-?' 2>&1 | Out-String)
    } catch {
        return $false
    }
    return [bool]($help -match '-f\s*<codepage>')
}

function Resolve-SqlCmdPath {
    if ($script:SqlCmdResolved) { return }

    $candidates = @(Get-Command sqlcmd -All -CommandType Application -ErrorAction SilentlyContinue)
    if ($candidates.Count -eq 0) {
        throw "sqlcmd.exe not found on PATH. Install SQL Server command-line tools (classic 'Command Line Utilities for SQL Server' or the newer 'sqlcmd' via winget) and try again."
    }

    $chosen = $candidates |
        Sort-Object { Get-SqlCmdFileVersion $_.Source } -Descending |
        Select-Object -First 1

    $script:SqlCmdPath = $chosen.Source
    $script:SqlCmdSupportsCodePage = Test-SqlCmdSupportsCodePage -Path $script:SqlCmdPath
    $script:SqlCmdResolved = $true

    $variant = if ($script:SqlCmdSupportsCodePage) { 'classic (ODBC-based)' } else { 'modern (go-sqlcmd)' }
    Write-Host "Using sqlcmd: $($script:SqlCmdPath) [$variant]" -ForegroundColor DarkGray
    if ($candidates.Count -gt 1) {
        Write-Host "  ($($candidates.Count) sqlcmd installations found on PATH; selected highest version)" -ForegroundColor DarkGray
    }
}

function Invoke-SaltoSqlCmd {
    param(
        [string]$SqlServer,
        [string]$Database,
        [string]$InputFile
    )

    Resolve-SqlCmdPath

    $previousOutputEncoding = [Console]::OutputEncoding
    $previousPsOutputEncoding = $OutputEncoding
    try {
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        [Console]::OutputEncoding = $utf8
        $OutputEncoding = $utf8
        $serverArg = Resolve-LocalDbServer -SqlServer $SqlServer
        $sqlCmdArgs = @('-S', $serverArg, '-E', '-d', $Database, '-W', '-s', '|', '-i', $InputFile)
        if ($script:SqlCmdSupportsCodePage) {
            $sqlCmdArgs += @('-f', 'i:65001,o:65001')
        }
        $raw = & $script:SqlCmdPath @sqlCmdArgs 2>&1
        if ($LASTEXITCODE -ne 0) {
            $msg = ($raw | Out-String).Trim()
            if ($msg) { throw "sqlcmd failed (exit $LASTEXITCODE): $msg" }
        }
        return @($raw)
    } finally {
        [Console]::OutputEncoding = $previousOutputEncoding
        $OutputEncoding = $previousPsOutputEncoding
    }
}

function Test-TcpPortOpen {
    param(
        [string]$HostName,
        [int]$Port,
        [int]$TimeoutMs = 3000
    )
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        $client.EndConnect($iar)
        return $client.Connected
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function Get-SqlLocalDbPath {
    $cmd = Get-Command SqlLocalDB.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) { return $cmd.Source }
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $root) { continue }
        $found = Get-ChildItem -Path (Join-Path $root 'Microsoft SQL Server\*\Tools\Binn\SqlLocalDB.exe') -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

function Resolve-LocalDbServer {
    # The modern go-sqlcmd cannot always turn "(localdb)\name" into a connection (seen
    # with a LocalDB 11.0 instance: "no named pipe instance matching 'name' returned
    # from host '(localdb)'"). LocalDB itself knows the instance's named pipe, and
    # both sqlcmd variants can connect to "np:\\.\pipe\LOCALDB#xxxxxxxx\tsql\query"
    # directly. The pipe name changes whenever the instance restarts, so it is looked
    # up on every run (and cached for the rest of that run). Anything that is not a
    # private LocalDB instance, or cannot be resolved, is returned unchanged.
    param([string]$SqlServer)

    if ($SqlServer -notmatch '^\(localdb\)\\(?<name>[^\\]+)$') { return $SqlServer }
    $name = $Matches['name'].Trim()
    if ($name.StartsWith('.')) { return $SqlServer }
    if ($script:LocalDbPipeCache.ContainsKey($name)) { return $script:LocalDbPipeCache[$name] }

    $resolved = $SqlServer
    $exe = Get-SqlLocalDbPath
    if ($exe) {
        $pattern = 'np:\\\\\.\\pipe\\LOCALDB#[0-9A-Fa-f]+\\tsql\\query'
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            for ($attempt = 0; $attempt -lt 2; $attempt++) {
                $info = @(& $exe info $name 2>&1 | ForEach-Object { "$_" })
                $m = [regex]::Match(($info -join "`n"), $pattern)
                if ($m.Success) { $resolved = $m.Value; break }
                if ($attempt -eq 0) { $null = @(& $exe start $name 2>&1) }
            }
        } finally {
            $ErrorActionPreference = $prevEap
        }
    }
    if ($resolved -ne $SqlServer) {
        Write-Host "[SQL] LocalDB instance '$name' resolved to its named pipe." -ForegroundColor DarkGray
    } else {
        Write-Host "[SQL] Could not resolve a named pipe for LocalDB instance '$name'; using '$SqlServer' as configured." -ForegroundColor DarkYellow
    }
    $script:LocalDbPipeCache[$name] = $resolved
    return $resolved
}

function Find-LocalDbInstanceProfiles {
    # A private LocalDB instance keeps its files in its owner's profile, so the
    # profile folders that contain Instances\<name> point at the owning account.
    param([string]$Name)

    $suffix = 'AppData\Local\Microsoft\Microsoft SQL Server Local DB\Instances\' + $Name
    $profiles = @()
    foreach ($root in @("$env:SystemDrive\Users", "$env:SystemRoot\ServiceProfiles")) {
        if (Test-Path -LiteralPath $root) {
            $profiles += @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
        }
    }
    $profiles += "$env:SystemRoot\System32\config\systemprofile"
    $found = @()
    foreach ($p in $profiles) {
        if (Test-Path -LiteralPath (Join-Path $p $suffix) -ErrorAction SilentlyContinue) { $found += $p }
    }
    return $found
}

function Test-IsElevated {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-RunningAsAccount {
    param([string]$Account)
    $current = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    return ($current -ieq $Account.Trim())
}

function ConvertTo-PsSingleQuoted([string]$Value) {
    return "'" + ($Value -replace "'", "''") + "'"
}

function Invoke-DelegatedRun {
    # Runs this script again as $Account (config "SqlRunAs") through a temporary
    # scheduled task and streams its output into this console. Needed when the Salto
    # database is a private LocalDB instance: only the account that owns the instance
    # (normally the Salto service account) can open it, and no interactive shell can
    # be started as such a service account. Requires an elevated shell. Returns the
    # exit code of the delegated run (0 = success, 1 = failed).
    param(
        [string]$Account,
        [hashtable]$ChildParams,
        [string]$ScriptFile
    )

    if (-not (Test-IsElevated)) {
        throw ("SqlRunAs is set to '$Account', so this run must be started as that account, which needs an elevated shell " +
            '(open PowerShell with "Run as administrator"). Alternatively run the script directly as that account ' +
            '(for example from a scheduled task) or add -NoDelegate to run in the current account anyway.')
    }

    $id = [guid]::NewGuid().ToString('N').Substring(0, 8)
    $dir = Join-Path $env:ProgramData "Sync-Salto-Unifi\delegate-$id"
    $taskName = "SaltoSyncDelegate-$id"
    $outFile = Join-Path $dir 'output.txt'
    $launcher = Join-Path $dir 'run.ps1'
    $marker = '__SALTO_DELEGATE_DONE__'
    $exitCode = 1
    $registered = $false
    $reader = $null
    try {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null

        # The delegated account must be able to read the launcher and write the output file.
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $grant = & icacls.exe $dir /grant ('{0}:(OI)(CI)M' -f $Account) /Q 2>&1
            $grantExit = $LASTEXITCODE
        } finally {
            $ErrorActionPreference = $prevEap
        }
        if ($grantExit -ne 0) {
            $grantText = @($grant | ForEach-Object { if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { [string]$_ } } |
                Where-Object { $_ -and $_.Trim() } | Select-Object -First 1)
            throw "Could not grant '$Account' access to the temporary run folder (does that account exist?): $($grantText -join ' ')"
        }

        $pairs = foreach ($key in $ChildParams.Keys) {
            $value = $ChildParams[$key]
            if ($value -is [bool]) { '{0} = ${1}' -f $key, $value.ToString().ToLower() }
            else { '{0} = {1}' -f $key, (ConvertTo-PsSingleQuoted ([string]$value)) }
        }
        $launcherText = @(
            "`$ErrorActionPreference = 'Continue'",
            ('$out = ' + (ConvertTo-PsSingleQuoted $outFile)),
            ('$p = @{ ' + ($pairs -join '; ') + ' }'),
            # One shared, auto-flushed writer: the parent reads the file while it is written.
            '$fs = [System.IO.File]::Open($out, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)',
            '$w = New-Object System.IO.StreamWriter($fs, (New-Object System.Text.UTF8Encoding($false)))',
            '$w.AutoFlush = $true',
            '$global:SaltoDelegateOk = $true',
            '& {',
            ('    try { & ' + (ConvertTo-PsSingleQuoted $ScriptFile) + ' @p }'),
            '    catch { $global:SaltoDelegateOk = $false; Write-Output ("[RUNAS] Delegated run failed: " + $_.Exception.Message) }',
            '} *>&1 | Out-String -Stream | ForEach-Object { $w.WriteLine([string]$_) }',
            '$code = if ($global:SaltoDelegateOk) { 0 } else { 1 }',
            ('$w.WriteLine("' + $marker + ' exit=" + $code)'),
            '$w.Dispose()'
        ) -join "`r`n"
        Set-Content -LiteralPath $launcher -Value $launcherText -Encoding UTF8

        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $launcher)
        $principal = New-ScheduledTaskPrincipal -UserId $Account -LogonType ServiceAccount
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2)
        try {
            Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Force | Out-Null
        } catch {
            throw "Could not create a scheduled task that runs as '$Account' (does that account exist, and is it a service account?): $($_.Exception.Message)"
        }
        $registered = $true

        Write-Host "[RUNAS] Running as '$Account' via temporary scheduled task '$taskName' (config SqlRunAs)..." -ForegroundColor DarkGray
        Start-ScheduledTask -TaskName $taskName

        $deadline = (Get-Date).AddMinutes(120)
        $done = $false
        $endedAt = $null
        while (-not $done -and (Get-Date) -lt $deadline) {
            if (-not $reader -and (Test-Path -LiteralPath $outFile)) {
                $stream = [System.IO.File]::Open($outFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $reader = New-Object System.IO.StreamReader($stream)
            }
            $gotLine = $false
            if ($reader) {
                while ($null -ne ($line = $reader.ReadLine())) {
                    $gotLine = $true
                    if ($line.StartsWith($marker)) {
                        if ($line -match 'exit=(\d+)') { $exitCode = [int]$Matches[1] }
                        $done = $true
                        break
                    }
                    Write-Host $line
                }
            }
            if ($done) { break }
            if ($gotLine) {
                $endedAt = $null
            } else {
                $state = [string](Get-ScheduledTask -TaskName $taskName).State
                if ($state -ne 'Running' -and $state -ne 'Queued') {
                    if (-not $endedAt) { $endedAt = Get-Date }
                    elseif (((Get-Date) - $endedAt).TotalSeconds -gt 20) {
                        $result = (Get-ScheduledTaskInfo -TaskName $taskName).LastTaskResult
                        throw ('The delegated run ended without reporting a result (task result 0x{0:X}). ' -f $result)
                    }
                } else {
                    $endedAt = $null
                }
            }
            Start-Sleep -Milliseconds 400
        }
        if (-not $done) { throw 'The delegated run did not finish within 120 minutes.' }
        Write-Host "[RUNAS] Delegated run finished (exit code $exitCode)." -ForegroundColor DarkGray
    } finally {
        if ($reader) { $reader.Dispose() }
        if ($registered) {
            Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        }
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
    return $exitCode
}

function Test-SqlServerPreflight {
    # Fails fast with an actionable message when the configured SqlServer cannot
    # possibly answer, instead of letting sqlcmd sit in a long, cryptic timeout.
    # Disable with "SqlPreflight": false in the config if it ever misjudges a setup.
    param($Cfg)

    if (-not (Get-ConfigBool $Cfg 'SqlPreflight' $true)) { return }

    $server = ([string]$Cfg.SqlServer).Trim()
    $who = "$env:USERDOMAIN\$env:USERNAME"
    $skipHint = ' Set "SqlPreflight": false in the config to skip this check.'

    # --- LocalDB: (localdb)\<instance> ---------------------------------------
    if ($server -match '^\(localdb\)\\(?<name>.+)$') {
        $name = $Matches['name'].Trim()
        if ($name.StartsWith('.\')) {
            Write-Host "[PREFLIGHT] LocalDB shared instance '$name' - instance check skipped." -ForegroundColor DarkGray
            return
        }
        $exe = Get-SqlLocalDbPath
        if (-not $exe) {
            Write-Host '[PREFLIGHT] SqlLocalDB.exe not found - LocalDB check skipped.' -ForegroundColor DarkYellow
            return
        }
        # "SqlLocalDB.exe i" lists instance names only (no localized labels to parse).
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $listed = @(& $exe i 2>&1 | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
        } finally {
            $ErrorActionPreference = $prevEap
        }
        if ($listed -notcontains $name) {
            $visible = if ($listed.Count -gt 0) { $listed -join ', ' } else { '(none)' }
            $ownerProfiles = @(Find-LocalDbInstanceProfiles -Name $name)
            $ownerText = if ($ownerProfiles.Count -gt 0) {
                " Instance '$name' has data in: " + ($ownerProfiles -join '; ') +
                ' (a folder under ServiceProfiles belongs to the service account of the same name, e.g. NT SERVICE\<service>; systemprofile is SYSTEM).'
            } else { '' }
            throw ("LocalDB instance '$name' is not visible to $who (visible to this account: $visible). " +
                'LocalDB instances belong to one Windows account: run this script as the account that owns the instance, ' +
                'normally the account the Salto ProAccess Space service runs as. Either run the script as that account (e.g. as a scheduled task), ' +
                'or set "SqlRunAs" in the config to that account so the script starts itself as it (needs an elevated shell).' + $ownerText +
                ' Find the service account with: ' +
                'Get-CimInstance Win32_Service | Where-Object { $_.PathName -like ''*salto*'' -or $_.DisplayName -like ''*salto*'' } | Select-Object Name, DisplayName, StartName, State.' + $skipHint)
        }
        Write-Host "[PREFLIGHT] LocalDB instance '$name' is visible to $who." -ForegroundColor DarkGray
        return
    }

    # --- host[\instance][,port] ----------------------------------------------
    if ($server -notmatch '^(?<host>[^\\,]+)(\\(?<inst>[^,]+))?(,(?<port>\d+))?$') {
        Write-Host "[PREFLIGHT] Unrecognised SqlServer format '$server' - check skipped." -ForegroundColor DarkYellow
        return
    }
    $hostName = ($Matches['host'] -replace '^(?i)tcp:', '').Trim()
    $instance = if ($Matches['inst']) { $Matches['inst'].Trim() } else { $null }
    $port = if ($Matches['port']) { [int]$Matches['port'] } else { $null }

    $localNames = @('.', '(local)', 'localhost', '127.0.0.1', '::1', $env:COMPUTERNAME)
    $isLocal = $localNames -contains $hostName

    # Local instance without explicit port: the SQL Server Windows service must exist and run.
    if ($isLocal -and -not $port) {
        $svcName = if ($instance) { 'MSSQL$' + $instance } else { 'MSSQLSERVER' }
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if (-not $svc) {
            $present = @(Get-Service -Name 'MSSQL*' -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            $presentText = if ($present.Count -gt 0) { $present -join ', ' } else { 'none' }
            throw ("SqlServer is '$server', but this machine has no SQL Server service '$svcName' (SQL services found: $presentText). " +
                "If the Salto database runs in LocalDB, set SqlServer to '(localdb)\<instance>' (see DBServerName in Salto's service.ini). " +
                'If it runs on another machine, set SqlServer to that host.' + $skipHint)
        }
        if ($svc.Status -ne 'Running') {
            throw ("SQL Server service '$svcName' is $($svc.Status), not Running. Start it with: Start-Service '$svcName'." + $skipHint)
        }
        Write-Host "[PREFLIGHT] SQL Server service '$svcName' is running." -ForegroundColor DarkGray
        return
    }

    # Remote host (or local with explicit port): resolve and test the TCP port.
    $target = if ($isLocal) { '127.0.0.1' } else { $hostName }
    if (-not $isLocal) {
        try {
            [void][System.Net.Dns]::GetHostAddresses($target)
        } catch {
            throw ("SQL Server host '$hostName' (from SqlServer '$server') cannot be resolved. Check the host name in the config." + $skipHint)
        }
    }
    $testPort = if ($port) { $port } elseif (-not $instance) { 1433 } else { $null }
    if ($testPort) {
        if (-not (Test-TcpPortOpen -HostName $target -Port $testPort)) {
            throw ("Cannot reach SQL Server at '$hostName' on TCP port $testPort (3 second timeout). " +
                'Check that the host is up, that SQL Server accepts TCP/IP connections, and that no firewall blocks the port.' + $skipHint)
        }
        Write-Host "[PREFLIGHT] SQL Server host '$hostName' answers on TCP port $testPort." -ForegroundColor DarkGray
    } else {
        Write-Host "[PREFLIGHT] Host '$hostName' resolves; named-instance port is found via SQL Browser (UDP 1434) - port check skipped." -ForegroundColor DarkGray
    }
}

function Read-Config([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config not found: $Path`nCopy unifi-sync-config.json.example to unifi-sync-config.json in the same folder as Sync-Salto-Unifi.ps1."
    }
    $raw = Get-Content -LiteralPath $Path -Raw
    try {
        $cfg = $raw | ConvertFrom-Json
    } catch {
        throw "Invalid JSON in config file: $Path`n$($_.Exception.Message)"
    }
    $required = @('UnifiHost', 'ApiTokenFile', 'SqlServer', 'Database', 'UserGroupName')
    foreach ($key in $required) {
        if (-not (Get-ObjProp $cfg $key)) { throw "Config missing required key: $key" }
    }
    if (-not (Test-Path -LiteralPath (Get-ObjProp $cfg 'ApiTokenFile'))) {
        throw "API token file not found: $(Get-ObjProp $cfg 'ApiTokenFile')"
    }
    $token = (Get-Content -LiteralPath (Get-ObjProp $cfg 'ApiTokenFile') -Raw).Trim()
    if (-not $token) { throw "API token file is empty: $($cfg.ApiTokenFile)" }
    $cfg | Add-Member -NotePropertyName Token -NotePropertyValue $token -Force
    Set-ConfigDefault $cfg 'OnlyActive' $true
    Set-ConfigDefault $cfg 'SaltoUserType' 'STAFF'
    Set-ConfigDefault $cfg 'RequireTag' $true
    Set-ConfigDefault $cfg 'PageSize' 200
    Set-ConfigDefault $cfg 'LogEnabled' $true
    Set-ConfigDefault $cfg 'LogDir' (Join-Path $PSScriptRoot 'logs\salto-unifi-sync')
    Set-ConfigDefault $cfg 'DeactivateWhenIneligible' $true
    Set-ConfigDefault $cfg 'DeleteOrphanNfcTokens' $true
    Set-ConfigDefault $cfg 'MaxDeactivationPercent' 10
    Set-ConfigDefault $cfg 'SqlPreflight' $true
    Set-ConfigDefault $cfg 'SqlRunAs' ''
    Set-ConfigDefault $cfg 'AutoUpdate' $false
    Set-ConfigDefault $cfg 'UpdateChannel' 'main'
    Set-ConfigDefault $cfg 'UpdateRepoOwner' 'jeroen-vermeulen'
    Set-ConfigDefault $cfg 'UpdateRepoName' 'salto-unifi-sync'
    Set-ConfigDefault $cfg 'UpdateCheckIntervalHours' 24
    return $cfg
}

function Get-DefaultLogDir {
    return (Join-Path $PSScriptRoot 'logs\salto-unifi-sync')
}

function Start-RunLogging {
    param(
        $Cfg,
        [string]$RunId
    )

    # Already logging (the delegating session starts it before the run): reuse it.
    if ($script:TranscriptActive) { return $script:RunLogPath }

    $enabled = $true
    $logDir = Get-DefaultLogDir
    if ($Cfg) {
        if ($Cfg.PSObject.Properties.Name -contains 'LogEnabled') {
            $enabled = [bool]$Cfg.LogEnabled
        }
        $configuredLogDir = Get-ObjProp $Cfg 'LogDir'
        if ($configuredLogDir) { $logDir = [string]$configuredLogDir }
    }
    if (-not $enabled) { return $null }

    $logPath = ''
    try {
        $logPath = Join-Path $logDir "salto-unifi-sync_$RunId.log"
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        }
        # Start-Transcript does not always fail when this account may not write to the
        # folder (the run then silently has no log), so probe the folder first.
        $probe = Join-Path $logDir ('.write-test-' + [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, 'x')
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        Start-Transcript -Path $logPath -Force -ErrorAction Stop | Out-Null
    } catch {
        $script:LogProblem = "Could not write a log file in '$logDir': $($_.Exception.Message)"
        Write-Host "[LOG] WARNING: $script:LogProblem" -ForegroundColor Yellow
        Write-Host '[LOG] The run continues without a log file. Set "LogDir" to a folder this account can write to.' -ForegroundColor Yellow
        return $null
    }
    $script:RunLogPath = $logPath
    $script:TranscriptActive = $true
    return $logPath
}

function Stop-RunLogging {
    if ($script:TranscriptActive) {
        try { Stop-Transcript | Out-Null } catch { }
        $script:TranscriptActive = $false
        if ($script:RunLogPath -and -not (Test-Path -LiteralPath $script:RunLogPath)) {
            Write-Host "[LOG] WARNING: the log file '$script:RunLogPath' does not exist after the run." -ForegroundColor Yellow
        }
    }
}

function Get-UpdateChannelLabel {
    param($Cfg)

    if (-not $Cfg) { return 'n/a' }
    $channel = [string](Get-ObjProp $Cfg 'UpdateChannel')
    if ([string]::IsNullOrWhiteSpace($channel)) { return 'main' }
    return $channel
}

function Write-RunBoundary {
    param(
        [ValidateSet('BEGIN', 'END')]
        [string]$Phase,

        [string]$Mode = '',
        [string]$Filter = '',
        [string]$UpdateChannel = '',
        [string]$LogPath = '',
        [string]$ConfigFile = '',
        [ValidateSet('SUCCESS', 'FAILED', '')]
        [string]$Status = '',
        [string]$Summary = '',
        $ErrorRecord = $null
    )

    $separator = '=' * 72
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $color = if ($Phase -eq 'BEGIN') { 'Cyan' } elseif ($Status -eq 'FAILED') { 'Red' } else { 'Green' }

    Write-Host ''
    Write-Host $separator -ForegroundColor $color
    if ($Phase -eq 'BEGIN') {
        Write-Host 'SALTO -> UNIFI SYNC RUN START' -ForegroundColor $color
        Write-Host "Timestamp : $timestamp"
        Write-Host "Version   : $ScriptVersion"
        if ($UpdateChannel) {
            Write-Host "Channel   : $UpdateChannel"
        }
        Write-Host "Mode      : $Mode"
        Write-Host "Filter    : $Filter"
        Write-Host "User      : $env:USERDOMAIN\$env:USERNAME"
        Write-Host "Computer  : $env:COMPUTERNAME"
        if ($ConfigFile) {
            Write-Host "Config    : $ConfigFile"
        }
        if ($LogPath) {
            Write-Host "Log file  : $LogPath"
        } elseif ($script:LogProblem) {
            Write-Host 'Log file  : (NOT written, see the [LOG] warning above)' -ForegroundColor Yellow
        } else {
            Write-Host 'Log file  : (logging disabled)'
        }
    } else {
        Write-Host 'SALTO -> UNIFI SYNC RUN END' -ForegroundColor $color
        Write-Host "Timestamp : $timestamp"
        Write-Host "Status    : $Status"
        if ($Summary) { Write-Host "Summary   : $Summary" }
        if ($ErrorRecord) {
            Write-Host "Error     : $($ErrorRecord.Exception.Message)"
            if ($ErrorRecord.ScriptStackTrace) {
                Write-Host 'Stack     :'
                Write-Host $ErrorRecord.ScriptStackTrace
            }
        }
        if ($script:RunLogPath) {
            Write-Host "Log file  : $script:RunLogPath"
        } elseif ($script:LogProblem) {
            Write-Host 'Log file  : (NOT written, see the [LOG] warning at the start)' -ForegroundColor Yellow
        }
    }
    Write-Host $separator -ForegroundColor $color
    Write-Host ''
}

function Parse-IdFilter([string]$FilterText) {
    $text = $FilterText.Trim()
    if ($text -eq 'ALL' -or $text -eq '*') {
        return @{ Type = 'All' }
    }
    if ($text -match ',') {
        $ids = @()
        foreach ($part in ($text -split ',')) {
            $piece = $part.Trim()
            if (-not $piece) { continue }
            if ($piece -notmatch '^\d+$') {
                throw "Invalid id in -Filter list: '$piece'"
            }
            $ids += [int]$piece
        }
        if ($ids.Count -eq 0) {
            throw "Invalid -Filter list '$FilterText': no ids found."
        }
        return @{ Type = 'List'; Ids = $ids }
    }
    if ($text -match '^\d+$') {
        $id = [int]$text
        return @{ Type = 'Range'; Min = $id; Max = $id }
    }
    if ($text -match '^(\d+)\s*-\s*(\d+)$') {
        $min = [int]$Matches[1]
        $max = [int]$Matches[2]
        if ($min -gt $max) { throw "Invalid filter range: $FilterText (min > max)" }
        return @{ Type = 'Range'; Min = $min; Max = $max }
    }
    throw "Invalid -Filter value '$FilterText'. Use ALL, a single id (e.g. 1000), a range (e.g. 10-20), or a quoted list (e.g. `"100,200,300`")."
}

function Test-IdInFilter {
    param(
        [Parameter(Mandatory)][int]$IdUser,
        [Parameter(Mandatory)]$IdFilter
    )
    switch ($IdFilter.Type) {
        'All' { return $true }
        'Range' { return ($IdUser -ge $IdFilter.Min -and $IdUser -le $IdFilter.Max) }
        'List' { return ($IdFilter.Ids -contains $IdUser) }
    }
    return $false
}

function Test-ScriptManagedUniFiUser {
    param(
        [Parameter(Mandatory)]$UniFiUserDetail
    )
    $empNo = [string](Get-ObjProp $UniFiUserDetail 'employee_number')
    $firstName = [string](Get-ObjProp $UniFiUserDetail 'first_name')
    if ($empNo -notmatch '^\d+$') {
        return @{ Ok = $false; Reason = 'No numeric employee_number' }
    }
    $expectedPrefix = '{0}_' -f $empNo
    if (-not $firstName.StartsWith($expectedPrefix, [System.StringComparison]::Ordinal)) {
        return @{ Ok = $false; Reason = "first_name does not start with '$expectedPrefix'" }
    }
    $underscoreIdx = $firstName.IndexOf('_')
    if ($underscoreIdx -lt 1) {
        return @{ Ok = $false; Reason = 'first_name missing id_user prefix' }
    }
    $parsedId = $firstName.Substring(0, $underscoreIdx)
    if ($parsedId -ne $empNo) {
        return @{ Ok = $false; Reason = 'first_name prefix does not match employee_number' }
    }
    return @{ Ok = $true; Reason = 'Script-managed naming' }
}

function Get-SaltoIneligibilityReason {
    param(
        $Cfg,
        $SaltoRecord
    )
    if ($null -eq $SaltoRecord) {
        return 'NOT_IN_SALTO'
    }
    if ($Cfg.SaltoUserType) {
        $userType = [string](Get-ObjProp $SaltoRecord 'UserType')
        if ($userType -ne $Cfg.SaltoUserType) {
            return 'SALTO_NOT_IN_SCOPE'
        }
    }
    $statusCode = Get-ObjProp $SaltoRecord 'StatusCode'
    if ($Cfg.OnlyActive -and $statusCode -ne 1) {
        if ($statusCode -eq 3) {
            return 'SALTO_DELETED'
        }
        return 'SALTO_INACTIVE'
    }
    $tagId = Normalize-SqlCmdValue ([string](Get-ObjProp $SaltoRecord 'TagID'))
    if ($Cfg.RequireTag -and [string]::IsNullOrWhiteSpace($tagId)) {
        return 'NO_SALTO_TAG'
    }
    return $null
}

function Get-IneligibleReasonText {
    param([string]$Reason)
    switch ($Reason) {
        'NO_SALTO_TAG' { return 'user has no active tag in Salto' }
        'SALTO_DELETED' { return 'user is deleted in Salto' }
        'NOT_IN_SALTO' { return 'user no longer exists in Salto' }
        'SALTO_NOT_IN_SCOPE' { return 'Salto user type is not in scope' }
        'SALTO_INACTIVE' { return 'user is not active in Salto' }
        default { return $Reason }
    }
}

function Format-PlanValue($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '(empty)' }
    return [string]$Value
}

function Write-SyncPlan {
    # Prints what a run would do: one block per changed user (action + old -> new per
    # field), then totals. -ShowAll also lists the users that need no change.
    param(
        [array]$Plan,
        [switch]$ShowAll
    )

    $sorted = @($Plan | Sort-Object id_user)
    $changed = @($sorted | Where-Object { $_.Action -notin @('NONE', 'SKIP_NON_BASIC') })
    $unchanged = @($sorted | Where-Object { $_.Action -in @('NONE', 'SKIP_NON_BASIC') })

    Write-Host ''
    Write-Host '=== Plan: users that will change ===' -ForegroundColor White
    if ($changed.Count -eq 0) {
        Write-Host '  (none - everything is in sync)' -ForegroundColor Green
    }
    foreach ($item in $changed) {
        $color = 'Yellow'
        if ($item.Action -match 'DEACTIVATE_USER|DELETE_USER') { $color = 'Red' }
        elseif ($item.Action -match 'CREATE_USER') { $color = 'Green' }
        Write-Host ('{0,6}  {1}' -f $item.id_user, $item.Name) -ForegroundColor $color -NoNewline
        Write-Host ('   [{0}]' -f $item.Action) -ForegroundColor DarkGray
        $lines = @($item.Changes)
        if ($lines.Count -eq 0) { $lines = @([string]$item.Details) }
        foreach ($line in $lines) {
            Write-Host ('          - {0}' -f $line)
        }
    }

    if ($ShowAll -and $unchanged.Count -gt 0) {
        Write-Host ''
        Write-Host '=== Plan: users without changes ===' -ForegroundColor White
        $unchanged | Format-Table id_user, Name, Status, SaltoTag, UniFiTag, Match, Details -AutoSize
    }

    Write-Host ''
    Write-Host '=== Plan: totals ===' -ForegroundColor White
    $counts = @{}
    foreach ($item in $changed) {
        foreach ($a in ($item.Action -split ',\s*')) {
            if (-not $counts.ContainsKey($a)) { $counts[$a] = 0 }
            $counts[$a]++
        }
    }
    foreach ($a in ($counts.Keys | Sort-Object)) {
        Write-Host ('  {0,-16} {1}' -f $a, $counts[$a])
    }
    Write-Host ('  {0,-16} {1}' -f 'no change', $unchanged.Count) -ForegroundColor DarkGray
    if (-not $ShowAll -and $unchanged.Count -gt 0) {
        Write-Host '  (add -ShowAll to list the users without changes)' -ForegroundColor DarkGray
    }
}

function Write-UniFiReconciliation {
    # Explains why "UniFi users loaded" and "Salto users in sync scope" differ: every
    # UniFi user is put in exactly one bucket. Only meaningful for -Filter ALL.
    param(
        [array]$Plan,
        [array]$UniFiUsers,
        [int]$SaltoCount
    )

    $matched = @{}
    $retire = @{}
    foreach ($item in $Plan) {
        if ($null -eq $item.UniFiUser) { continue }
        $uid = [string](Get-ObjProp $item.UniFiUser 'id')
        if (-not $uid) { continue }
        if ($item.Match -eq 'script_managed') { $retire[$uid] = $item } else { $matched[$uid] = $item }
    }
    $createCount = @($Plan | Where-Object { $_.Action -match 'CREATE_USER' }).Count

    $untouched = @()
    foreach ($u in $UniFiUsers) {
        $uid = [string](Get-ObjProp $u 'id')
        if ($matched.ContainsKey($uid) -or $retire.ContainsKey($uid)) { continue }
        $managed = Test-ScriptManagedUniFiUser -UniFiUserDetail $u
        if (-not $managed.Ok) {
            $reason = 'not script-managed: ' + $managed.Reason
        } elseif ((Get-ObjProp $u 'status') -ne 'ACTIVE') {
            $reason = 'script-managed, already inactive in UniFi'
        } else {
            $reason = 'script-managed and active, but no action planned (non-basic user, or still eligible in Salto)'
        }
        $untouched += [pscustomobject]@{
            Name   = ('{0} {1}' -f (Get-ObjProp $u 'first_name'), (Get-ObjProp $u 'last_name')).Trim()
            Status = [string](Get-ObjProp $u 'status')
            Reason = $reason
        }
    }

    Write-Host ''
    Write-Host '=== UniFi versus Salto: where do the numbers come from ===' -ForegroundColor White
    Write-Host ('  UniFi users loaded                                  {0,4}' -f @($UniFiUsers).Count)
    Write-Host ('    matched with a Salto user in scope                {0,4}' -f $matched.Count)
    Write-Host ('    script-managed, no longer in Salto scope (change) {0,4}' -f $retire.Count)
    Write-Host ('    left alone by the script                          {0,4}' -f $untouched.Count)
    Write-Host ('  Salto users in scope                                {0,4}' -f $SaltoCount)
    Write-Host ('    not yet in UniFi (will be created)                {0,4}' -f $createCount)
    if ($untouched.Count -gt 0) {
        Write-Host ''
        Write-Host '  UniFi users left alone by the script:' -ForegroundColor White
        $untouched | Sort-Object Reason, Name | Select-Object -First 60 | ForEach-Object {
            Write-Host ('    {0,-34} {1,-12} {2}' -f $_.Name, $_.Status, $_.Reason)
        }
        if ($untouched.Count -gt 60) { Write-Host ('    ... and {0} more' -f ($untouched.Count - 60)) -ForegroundColor DarkGray }
    }
}

function Get-IneligibleAction {
    param([string]$Reason)
    switch ($Reason) {
        'NO_SALTO_TAG' { return 'DELETE_USER' }
        'SALTO_DELETED' { return 'DELETE_USER' }
        'NOT_IN_SALTO' { return 'DELETE_USER' }
        default { return 'DEACTIVATE_USER' }
    }
}

function New-CurlAuthConfigFile {
    param(
        [Parameter(Mandatory)][string]$Token
    )
    # Passing the bearer token via -K (a curl "config" file) keeps it out of the
    # process command line, where any other local process/user could otherwise
    # read it (e.g. via Get-CimInstance Win32_Process).
    $path = Join-Path $env:TEMP ("unifi-auth-{0}.curlcfg" -f [guid]::NewGuid().ToString('N'))
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $escaped = $Token.Replace('"', '\"')
    [System.IO.File]::WriteAllText($path, "header = ""Authorization: Bearer $escaped""`n", $utf8)
    return $path
}

function Invoke-UniFiJson {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Url,
        [string]$BodyPath = $null,
        [string]$Token
    )
    $respFile = Join-Path $env:TEMP ("unifi-resp-{0}.json" -f [guid]::NewGuid().ToString('N'))
    $authCfg = New-CurlAuthConfigFile -Token $Token
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    try {
        $args = @(
            '-sk', '-m', '120',
            '-K', $authCfg,
            '-X', $Method,
            $Url,
            '-H', 'accept: application/json',
            '-o', $respFile
        )
        if ($BodyPath) {
            $args += @('-H', 'content-type: application/json; charset=utf-8', '--data-binary', "@$BodyPath")
        }
        $stderr = & curl.exe @args 2>&1
        if ($LASTEXITCODE -ne 0) {
            $errMsg = ($stderr | Out-String).Trim()
            throw "curl failed (exit $LASTEXITCODE) for $Method $Url`: $errMsg"
        }
        if (-not (Test-Path -LiteralPath $respFile)) {
            throw "Empty response from $Url (no output file)"
        }
        $rawText = [System.IO.File]::ReadAllText($respFile, $utf8).Trim()
        if (-not $rawText) { throw "Empty response from $Url" }
        $json = $rawText | ConvertFrom-Json
        $code = Get-ObjProp $json 'code'
        if ($code -and $code -ne 'SUCCESS') {
            $msg = Get-ObjProp $json 'msg'
            throw "$Method $Url failed: $code $msg"
        }
        return $json
    } finally {
        Remove-Item -LiteralPath $respFile -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $authCfg -ErrorAction SilentlyContinue
    }
}

function Get-AllUniFiUsers($Cfg) {
    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $page = 1
    $all = @()
    $pageSize = [int]$Cfg.PageSize
    if ($pageSize -lt 1) { $pageSize = 200 }

    while ($true) {
        $resp = Invoke-UniFiJson -Method GET -Url "$api/users?page_num=$page&page_size=$pageSize" -Token $Cfg.Token
        $batch = @()
        $batchData = Get-ObjProp $resp 'data'
        if ($batchData) { $batch = @($batchData) }
        if ($batch.Count -eq 0) { break }

        $all += $batch
        if ($batch.Count -lt $pageSize) { break }
        $page++
    }

    return $all
}

function Test-BasicManagedUser {
    param(
        [Parameter(Mandatory)]$UserDetail,
        [Parameter(Mandatory)]$IdentityAssignments
    )
    $u = $UserDetail
    if (Get-ObjProp $u 'username') { return @{ Ok = $false; Reason = 'Has username (not a basic local user)' } }
    if ((Get-ObjProp $u 'user_email') -or (Get-ObjProp $u 'email')) { return @{ Ok = $false; Reason = 'Has email (not a basic local user)' } }
    if ($IdentityAssignments) {
        $identityData = Get-ObjProp $IdentityAssignments 'data'
        if ($identityData) {
        foreach ($prop in $identityData.PSObject.Properties) {
            $val = $prop.Value
            if ($null -eq $val) { continue }
            if ($val -is [System.Array] -and $val.Count -gt 0) { return @{ Ok = $false; Reason = 'Has UniFi Identity assignments' } }
            if ($val -is [psobject]) {
                $hasItems = $false
                foreach ($inner in $val.PSObject.Properties) {
                    if ($inner.Value -and @($inner.Value).Count -gt 0) { $hasItems = $true; break }
                }
                if ($hasItems) { return @{ Ok = $false; Reason = 'Has UniFi Identity assignments' } }
            }
        }
        }
    }
    return @{ Ok = $true; Reason = 'Basic local user' }
}

function Get-SaltoUsers {
    param(
        $Cfg,
        $IdFilter,
        [switch]$LookupMode
    )

    $where = @()
    if (-not $LookupMode) {
        if ($Cfg.OnlyActive) { $where += 'u.status = 1' }
        if ($Cfg.SaltoUserType) { $where += "ut.description = N'$($Cfg.SaltoUserType.Replace("'", "''"))'" }
        if ($Cfg.RequireTag) { $where += 'tag.ROMCode IS NOT NULL' }
    }
    if ($IdFilter.Type -eq 'Range') {
        $where += "u.id_user BETWEEN $($IdFilter.Min) AND $($IdFilter.Max)"
    } elseif ($IdFilter.Type -eq 'List') {
        $idList = ($IdFilter.Ids | ForEach-Object { [string][int]$_ }) -join ','
        $where += "u.id_user IN ($idList)"
    }
    $whereSql = if ($where.Count -gt 0) { 'WHERE ' + ($where -join ' AND ') } else { '' }

    $query = @"
SELECT
    u.id_user,
    u.FirstName,
    u.LastName,
    u.status AS StatusCode,
    us.description AS StatusText,
    ut.description AS UserType,
    UPPER(tag.ROMCode) AS TagID
FROM tb_Users u
LEFT JOIN lt_UserStatus us ON u.status = us.code
LEFT JOIN lt_UserType ut ON u.type = ut.code
OUTER APPLY (
    SELECT TOP 1 c.ROMCode
    FROM tb_Cards c
    WHERE c.id_user = u.id_user
      AND c.CancellationDateUTC IS NULL
      AND (c.ExpirationDate IS NULL OR c.ExpirationDate > GETUTCDATE())
      AND c.ROMCode IS NOT NULL AND LTRIM(RTRIM(c.ROMCode)) <> ''
    ORDER BY c.Cardcode
) tag
$whereSql
ORDER BY u.id_user;
"@

    $tmpSql = Join-Path $env:TEMP ('salto-unifi-sync-query{0}.sql' -f $(if ($LookupMode) { '-lookup' } else { '' }))
    Write-JsonNoBom $tmpSql $query
    $raw = Invoke-SaltoSqlCmd -SqlServer $Cfg.SqlServer -Database $Cfg.Database -InputFile $tmpSql
    if (-not $raw) { throw 'sqlcmd returned no data. Is SQL Server reachable for the current Windows user?' }

    $rows = @()
    foreach ($line in $raw) {
        if (-not $line -or $line -match '^-+$' -or $line -match 'rows affected' -or $line -match '^id_user\|') { continue }
        $parts = $line -split '\|', -1
        if ($parts.Count -lt 7) { continue }
        if (-not ($parts[0] -match '^\d+$')) { continue }
        $tagId = Normalize-SqlCmdValue $parts[6]
        $rows += [pscustomobject]@{
            id_user    = [int]$parts[0]
            FirstName  = Normalize-SqlCmdValue $parts[1]
            LastName   = Normalize-SqlCmdValue $parts[2]
            StatusCode = if ($parts[3] -match '^\d+$') { [int]$parts[3] } else { 0 }
            StatusText = Normalize-SqlCmdValue $parts[4]
            UserType   = Normalize-SqlCmdValue $parts[5]
            TagID      = $tagId
        }
    }
    return $rows
}

function Get-UniFiFirstName([int]$IdUser, [string]$SaltoFirstName) {
    $voornaam = if ($SaltoFirstName) { $SaltoFirstName.Trim() } else { 'Unknown' }
    return "{0}_{1}" -f $IdUser, $voornaam
}

function Get-DesiredState([pscustomobject]$SaltoUser) {
    return [pscustomobject]@{
        id_user         = $SaltoUser.id_user
        employee_number = [string]$SaltoUser.id_user
        first_name      = Get-UniFiFirstName $SaltoUser.id_user $SaltoUser.FirstName
        last_name       = $SaltoUser.LastName.Trim()
        tag_id          = $SaltoUser.TagID.ToUpper()
        status_text     = $SaltoUser.StatusText
    }
}

function Format-SyncUserLabel($Desired) {
    return '{0} {1}' -f $Desired.first_name, $Desired.last_name
}

function Build-NfcTokenMap {
    param($Cfg)

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $map = @{}
    $page = 1
    $pageSize = 500

    do {
        $resp = Invoke-UniFiJson -Method GET -Url "$api/credentials/nfc_cards/tokens?page_num=$page&page_size=$pageSize" -Token $Cfg.Token
        $cardData = Get-ObjProp $resp 'data'
        if ($cardData) {
            foreach ($card in @($cardData)) {
                $token = Get-ObjProp $card 'token'
                $nfcId = Get-ObjProp $card 'nfc_id'
                $note = Get-ObjProp $card 'note'
                $aliasVal = Get-ObjProp $card 'alias'
                $uid = $null
                if ($nfcId) {
                    $uid = [string]$nfcId.ToUpper()
                } elseif ($note -and ([string]$note).ToUpper() -match '^[0-9A-F]{14}$') {
                    $uid = [string]$note.ToUpper()
                }
                if ($uid) {
                    $map[$uid] = $uid
                    if ($token) { $map[[string]$token] = $uid }
                }
                if ($token -and $aliasVal) {
                    $map["token_alias:$token"] = [string]$aliasVal
                    if ($aliasVal -match '^salto-\d+$') {
                        $map["alias:$aliasVal"] = [string]$token
                    }
                    if ($uid) {
                        $map[[string]$aliasVal] = $uid
                    }
                }
            }
        }

        $pagination = Get-ObjProp $resp 'pagination'
        $total = 0
        if ($pagination) {
            $total = [int](Get-ObjProp $pagination 'total')
        }
        $page++
    } while ($cardData -and (($page - 1) * $pageSize) -lt $total)

    return $map
}

function Get-UserPhysicalNfcTag {
    param(
        $UniFiUserDetail,
        $NfcTokenMap = $null
    )

    $cards = Get-ObjProp $UniFiUserDetail 'nfc_cards'
    if (-not $cards) { return $null }
    foreach ($card in @($cards)) {
        foreach ($field in @('nfc_id', 'uid', 'nfc_uid', 'serial_number')) {
            $val = Get-ObjProp $card $field
            if ($val) { return [string]$val.ToUpper() }
        }
        $token = Get-ObjProp $card 'token'
        if ($token -and $NfcTokenMap -and $NfcTokenMap.ContainsKey([string]$token)) {
            return [string]$NfcTokenMap[[string]$token].ToUpper()
        }
    }
    return $null
}

function Get-UserNfcUid {
    param(
        $UniFiUserDetail,
        $NfcTokenMap = $null
    )

    $cards = Get-ObjProp $UniFiUserDetail 'nfc_cards'
    if (-not $cards) { return $null }
    foreach ($card in @($cards)) {
        foreach ($field in @('nfc_id', 'uid', 'nfc_uid', 'serial_number')) {
            $val = Get-ObjProp $card $field
            if ($val) { return [string]$val.ToUpper() }
        }
        $token = Get-ObjProp $card 'token'
        if ($token -and $NfcTokenMap -and $NfcTokenMap.ContainsKey([string]$token)) {
            return $NfcTokenMap[[string]$token]
        }
    }
    return $null
}

function Test-NfcInSync {
    param(
        $UniFiUserDetail,
        $Desired,
        $NfcTokenMap
    )

    $uid = Get-UserNfcUid $UniFiUserDetail -NfcTokenMap $NfcTokenMap
    if ($uid -eq $Desired.tag_id) { return $true }

    # Fallback for when the card's own UID couldn't be resolved directly: the
    # "salto-{id_user}" alias is per-user, not per-tag, and stays attached to
    # whichever token was last assigned - merely owning that alias slot does NOT
    # prove the tag is current (e.g. after a physical card replacement). Only
    # trust it if the UID currently behind that alias actually matches the
    # desired tag.
    if ($NfcTokenMap) {
        $expectedAlias = "salto-$($Desired.id_user)"
        if ($NfcTokenMap.ContainsKey($expectedAlias) -and $NfcTokenMap[$expectedAlias] -eq $Desired.tag_id) {
            return $true
        }
    }

    return $false
}

function Get-UserNfcDisplayTag {
    param(
        $UniFiUserDetail,
        $Desired,
        $NfcTokenMap
    )

    $uid = Get-UserNfcUid $UniFiUserDetail -NfcTokenMap $NfcTokenMap
    if ($uid) { return $uid }
    if (Test-NfcInSync $UniFiUserDetail $Desired $NfcTokenMap) {
        return $Desired.tag_id
    }
    return $null
}

function Build-UniFiUserIndex {
    param(
        $Cfg,
        [array]$UniFiUsers
    )

    # The bulk user list (Get-AllUniFiUsers) already returns every field this index
    # needs - employee_number, first_name, last_name, status, username, email,
    # user_email, nfc_cards - identical in content to the per-user detail endpoint
    # (verified against the live API). Fetching /users/$id again per user here used
    # to cost one curl.exe process + TLS handshake per UniFi user for no new data.
    $nfcTokenMap = Build-NfcTokenMap -Cfg $Cfg
    $index = @{
        ByEmployee   = @{}
        ByName       = @{}
        ByTag        = @{}
        NfcTokenMap  = $nfcTokenMap
    }

    foreach ($u in $UniFiUsers) {
        $userId = Get-ObjProp $u 'id'
        if (-not $userId) { continue }

        $empNo = Get-ObjProp $u 'employee_number'
        if ($empNo -match '^\d+$') {
            $index.ByEmployee[[string]$empNo] = $u
        }

        $firstName = Get-ObjProp $u 'first_name'
        $lastName = Get-ObjProp $u 'last_name'
        if ($firstName -and $lastName) {
            $index.ByName["$firstName|$lastName"] = $u
        }

        $tag = Get-UserNfcUid $u -NfcTokenMap $nfcTokenMap
        if ($tag) {
            $index.ByTag[$tag] = $u
        }
    }

    return $index
}

function Find-MatchedUniFiUser {
    param(
        $Desired,
        $SaltoUser,
        $Index
    )

    $emp = [string]$Desired.id_user
    if ($Index.ByEmployee.ContainsKey($emp)) {
        return @{ User = $Index.ByEmployee[$emp]; MatchBy = 'employee_number' }
    }

    $expectedKey = "$($Desired.first_name)|$($Desired.last_name)"
    if ($Index.ByName.ContainsKey($expectedKey)) {
        return @{ User = $Index.ByName[$expectedKey]; MatchBy = 'expected_name' }
    }

    $legacyKey = "$($SaltoUser.FirstName)|$($SaltoUser.LastName)"
    if ($Index.ByName.ContainsKey($legacyKey)) {
        return @{ User = $Index.ByName[$legacyKey]; MatchBy = 'legacy_name' }
    }

    if ($Index.ByTag.ContainsKey($Desired.tag_id)) {
        return @{ User = $Index.ByTag[$Desired.tag_id]; MatchBy = 'nfc_tag' }
    }

    return @{ User = $null; MatchBy = 'none' }
}

function Get-UniFiGroupMemberIds {
    param(
        [string]$GroupId,
        $Cfg
    )
    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $resp = Invoke-UniFiJson -Method GET -Url "$api/user_groups/$GroupId/users" -Token $Cfg.Token
    $data = Get-ObjProp $resp 'data'
    $set = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($member in @($data)) {
        $id = Get-ObjProp $member 'id'
        if ($id) { [void]$set.Add([string]$id) }
    }
    return $set
}

function Test-UserInGroup {
    param(
        [string]$UserId,
        [System.Collections.Generic.HashSet[string]]$GroupMemberIds
    )
    if (-not $GroupMemberIds) { return $false }
    return $GroupMemberIds.Contains($UserId)
}

function Add-DeactivationPlanItems {
    param(
        [array]$Plan,
        $Cfg,
        $IdFilter,
        [array]$SaltoUsers,
        [array]$SaltoLookup,
        $UniIndex,
        $NfcTokenMap,
        [string]$Api,
        [hashtable]$ClaimedUniFiIds = @{}
    )

    if (-not (Get-ConfigBool $Cfg 'DeactivateWhenIneligible' $true)) {
        return ,@($Plan)
    }

    $eligibleIds = @{}
    foreach ($salto in $SaltoUsers) {
        $eligibleIds[[string]$salto.id_user] = $true
    }

    $lookupById = @{}
    foreach ($rec in $SaltoLookup) {
        $lookupById[[string]$rec.id_user] = $rec
    }

    foreach ($empNo in @($UniIndex.ByEmployee.Keys)) {
        if (-not (Test-IdInFilter -IdUser ([int]$empNo) -IdFilter $IdFilter)) { continue }
        if ($eligibleIds.ContainsKey($empNo)) { continue }

        $detail = $UniIndex.ByEmployee[$empNo]
        # A UniFi user that an eligible Salto user already matched (for example the same
        # person re-created in Salto under a new id_user, with the tag moved along) is
        # updated through that match; deleting it here would make the update fail later.
        if ($ClaimedUniFiIds.ContainsKey([string](Get-ObjProp $detail 'id'))) { continue }
        if ((Get-ObjProp $detail 'status') -ne 'ACTIVE') { continue }

        $managed = Test-ScriptManagedUniFiUser -UniFiUserDetail $detail
        if (-not $managed.Ok) { continue }

        $uniUserId = Get-ObjProp $detail 'id'
        $identity = Invoke-UniFiJson -Method GET -Url "$Api/users/$uniUserId/identity/assignments" -Token $Cfg.Token
        $basic = Test-BasicManagedUser -UserDetail $detail -IdentityAssignments $identity
        if (-not $basic.Ok) { continue }

        $saltoRec = $null
        if ($lookupById.ContainsKey($empNo)) { $saltoRec = $lookupById[$empNo] }
        $reason = Get-SaltoIneligibilityReason -Cfg $Cfg -SaltoRecord $saltoRec
        if (-not $reason) { continue }

        $firstName = [string](Get-ObjProp $detail 'first_name')
        $lastName = [string](Get-ObjProp $detail 'last_name')
        $idUser = [int]$empNo
        $desired = [pscustomobject]@{
            id_user         = $idUser
            employee_number = $empNo
            first_name      = $firstName
            last_name       = $lastName
            tag_id          = if ($saltoRec -and $saltoRec.TagID) { $saltoRec.TagID } else { '' }
            status_text     = if ($saltoRec) { $saltoRec.StatusText } else { '-' }
        }

        $currentTag = Get-UserNfcDisplayTag $detail $desired $NfcTokenMap
        $Plan += [pscustomobject]@{
            id_user   = $idUser
            Name      = "$firstName $lastName"
            SaltoTag  = if ($saltoRec -and $saltoRec.TagID) { $saltoRec.TagID } else { '-' }
            UniFiTag  = if ($currentTag) { $currentTag } else { '-' }
            Match     = 'script_managed'
            Status    = $desired.status_text
            Action    = (Get-IneligibleAction $reason)
            Details   = $reason
            Changes   = @("$(if ((Get-IneligibleAction $reason) -eq 'DELETE_USER') { 'delete user from UniFi' } else { 'deactivate user in UniFi' }) ($(Get-IneligibleReasonText $reason))")
            Desired   = $desired
            UniFiUser = $detail
        }
    }

    return ,@($Plan)
}

function Build-SyncPlan {
    param(
        [array]$SaltoUsers,
        [array]$SaltoLookup,
        [array]$UniFiUsers,
        $Cfg,
        [string]$SyncMode,
        $IdFilter
    )

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $groupResp = Invoke-UniFiJson -Method GET -Url "$api/user_groups" -Token $Cfg.Token
    $groupData = Get-ObjProp $groupResp 'data'
    $group = $groupData | Where-Object { (Get-ObjProp $_ 'name') -eq $Cfg.UserGroupName } | Select-Object -First 1
    if (-not $group) {
        throw "UniFi user group '$($Cfg.UserGroupName)' not found."
    }
    $groupId = Get-ObjProp $group 'id'
    if (-not $groupId) {
        throw "UniFi user group '$($Cfg.UserGroupName)' has no id."
    }

    Write-Host 'Indexing UniFi users (name, employee_number, NFC tag)...'
    $uniIndex = Build-UniFiUserIndex -Cfg $Cfg -UniFiUsers $UniFiUsers
    $nfcTokenMap = $uniIndex.NfcTokenMap
    $groupMemberIds = Get-UniFiGroupMemberIds -GroupId $groupId -Cfg $Cfg

    $plan = @()
    foreach ($salto in $SaltoUsers) {
        $desired = Get-DesiredState $salto
        $actions = @()
        $changes = @()
        $match = Find-MatchedUniFiUser -Desired $desired -SaltoUser $salto -Index $uniIndex
        $uniDetail = $match.User
        $uniUser = $null
        $matchBy = $match.MatchBy
        $currentTag = $null
        $uniSummary = '(not in UniFi)'

        if ($uniDetail) {
            $uniUser = $uniDetail
            $uniUserId = Get-ObjProp $uniDetail 'id'
            $identity = Invoke-UniFiJson -Method GET -Url "$api/users/$uniUserId/identity/assignments" -Token $Cfg.Token
            $basic = Test-BasicManagedUser -UserDetail $uniDetail -IdentityAssignments $identity
            if (-not $basic.Ok) {
                $legacyTag = Get-UserPhysicalNfcTag -UniFiUserDetail $uniDetail -NfcTokenMap $nfcTokenMap
                $legacySummary = "$(Get-ObjProp $uniDetail 'first_name') $(Get-ObjProp $uniDetail 'last_name')"
                if ($legacyTag -and $legacyTag -eq $desired.tag_id) {
                    $plan += [pscustomobject]@{
                        id_user   = $desired.id_user
                        Name      = (Format-SyncUserLabel $desired)
                        SaltoTag  = $desired.tag_id
                        UniFiTag  = $legacyTag
                        Match     = $matchBy
                        Status    = $desired.status_text
                        Action    = 'NONE'
                        Details   = "Non-basic user; Salto tag already correct ($legacySummary)"
                        Changes   = @()
                        Desired   = $desired
                        UniFiUser = $uniUser
                        GroupId   = $groupId
                    }
                    continue
                }

                $currentTag = if ($legacyTag) { $legacyTag } else { '-' }
                $tagNote = if ($legacyTag) { "wrong tag $legacyTag" } else { 'no tag' }
                $uniSummary = "Non-basic ($($basic.Reason)); legacy $legacySummary ($tagNote) -> new script user"
                $matchBy = 'legacy_non_basic'
                $uniDetail = $null
                $uniUser = $null
            }
        }

        if ($uniDetail) {
            $uniUserId = Get-ObjProp $uniDetail 'id'
            $uniSummary = "$(Get-ObjProp $uniDetail 'first_name') $(Get-ObjProp $uniDetail 'last_name')"
            $currentTag = Get-UserNfcDisplayTag $uniDetail $desired $nfcTokenMap
            $inGroup = Test-UserInGroup -UserId $uniUserId -GroupMemberIds $groupMemberIds

            $nameDiff = Test-NameDiffers `
                -UniFirst (Get-ObjProp $uniDetail 'first_name') `
                -UniLast (Get-ObjProp $uniDetail 'last_name') `
                -DesiredFirst $desired.first_name `
                -DesiredLast $desired.last_name
            $empDiff = ((Get-ObjProp $uniDetail 'employee_number') -ne $desired.employee_number)
            $tagDiff = -not (Test-NfcInSync $uniDetail $desired $nfcTokenMap)
            $groupDiff = -not $inGroup
            $statusDiff = ((Get-ObjProp $uniDetail 'status') -ne 'ACTIVE')

            if ($nameDiff) {
                $changes += ("name: '{0} {1}' -> '{2} {3}'" -f (Get-ObjProp $uniDetail 'first_name'), (Get-ObjProp $uniDetail 'last_name'), $desired.first_name, $desired.last_name)
            }
            if ($empDiff) {
                $changes += ("employee_number: '{0}' -> '{1}'" -f (Format-PlanValue (Get-ObjProp $uniDetail 'employee_number')), $desired.employee_number)
            }
            if ($tagDiff) {
                $changes += ('NFC tag: {0} -> {1}' -f (Format-PlanValue $currentTag), $desired.tag_id)
            }
            if ($groupDiff) {
                $changes += ("add to user group '{0}'" -f $Cfg.UserGroupName)
            }
            if ($statusDiff) {
                $changes += ('status: {0} -> ACTIVE' -f (Format-PlanValue (Get-ObjProp $uniDetail 'status')))
            }

            if ($SyncMode -eq 'FullSync') {
                if ($nameDiff -or $empDiff) { $actions += 'UPDATE_USER' }
                if ($tagDiff) { $actions += 'UPDATE_NFC' }
                if ($groupDiff) { $actions += 'ADD_GROUP' }
                if ($statusDiff) { $actions += 'ACTIVATE_USER' }
                if ($actions.Count -eq 0) {
                    $actions += 'REFRESH'
                    $changes += 'no difference found; FullSync re-writes all mapped fields'
                }
            } else {
                if ($nameDiff -or $empDiff) { $actions += 'UPDATE_USER' }
                if ($tagDiff) { $actions += 'UPDATE_NFC' }
                if ($groupDiff) { $actions += 'ADD_GROUP' }
                if ($statusDiff) { $actions += 'ACTIVATE_USER' }
            }
            if ($actions.Count -eq 0) { $actions += 'NONE' }
        } else {
            $actions += 'CREATE_USER'
            $actions += 'ADD_GROUP'
            $actions += 'ASSIGN_NFC'
            $changes += ("create user '{0}' (employee_number {1})" -f (Format-SyncUserLabel $desired), $desired.employee_number)
            $changes += ("add to user group '{0}'" -f $Cfg.UserGroupName)
            $changes += ('assign NFC tag {0}' -f $desired.tag_id)
        }

        $plan += [pscustomobject]@{
            id_user   = $desired.id_user
            Name      = (Format-SyncUserLabel $desired)
            SaltoTag  = $desired.tag_id
            UniFiTag  = if ($currentTag) { $currentTag } else { '-' }
            Match     = $matchBy
            Status    = $desired.status_text
            Action    = ($actions -join ', ')
            Details   = "UniFi: $uniSummary"
            Changes   = @($changes)
            Desired   = $desired
            UniFiUser = $uniUser
            GroupId   = $groupId
        }
    }

    $claimedUniFiIds = @{}
    foreach ($planItem in $plan) {
        if ($planItem.UniFiUser) {
            $claimedId = [string](Get-ObjProp $planItem.UniFiUser 'id')
            if ($claimedId) { $claimedUniFiIds[$claimedId] = $true }
        }
    }

    $plan = Add-DeactivationPlanItems `
        -Plan $plan `
        -Cfg $Cfg `
        -IdFilter $IdFilter `
        -SaltoUsers $SaltoUsers `
        -SaltoLookup $SaltoLookup `
        -UniIndex $uniIndex `
        -NfcTokenMap $nfcTokenMap `
        -Api $api `
        -ClaimedUniFiIds $claimedUniFiIds

    return ,@($plan)
}

function Get-FirstArrayItem($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return $null }
        return $Value[0]
    }
    return $Value
}

function Remove-StaleSaltoAliasToken {
    param(
        [string]$Token,
        [string]$UserId,
        $Cfg
    )
    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    if ($UserId) {
        Unassign-NfcTokensFromUser -UserId $UserId -Tokens @($Token) -Cfg $Cfg -UserLabel 'previous NFC tag'
    }
    if (-not (Test-NfcTokenExists -Token $Token -Cfg $Cfg)) { return }
    try {
        Invoke-UniFiJson -Method DELETE -Url "$api/credentials/nfc_cards/tokens/$Token" -Token $Cfg.Token | Out-Null
        Write-Host "[NFC DELETE] Removed superseded token from inventory" -ForegroundColor Magenta
    } catch {
        Write-Host "[NFC DELETE FAIL] Could not delete superseded token: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Ensure-NfcToken {
    param(
        [string]$TagId,
        [int]$SaltoUserId,
        $Cfg,
        [string]$UserId = $null
    )
    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $alias = "salto-$SaltoUserId"
    $cards = Invoke-UniFiJson -Method GET -Url "$api/credentials/nfc_cards/tokens?page_size=500" -Token $Cfg.Token
    $cardData = Get-ObjProp $cards 'data'

    # A token already imported for this exact tag wins outright, regardless of alias.
    $existing = $cardData | Where-Object {
        $nfcId = Get-ObjProp $_ 'nfc_id'
        $note = Get-ObjProp $_ 'note'
        ($nfcId -and ([string]$nfcId).ToUpper() -eq $TagId) -or ($note -and ([string]$note).ToUpper() -eq $TagId)
    } | Select-Object -First 1
    if ($existing) {
        $token = Get-ObjProp $existing 'token'
        if ($token) { return $token }
    }

    # No token for this tag yet. If this user's alias is still attached to an older
    # (superseded) token - e.g. after a physical card replacement - clean that up
    # first so "salto-{id_user}" always resolves to exactly one, current token.
    $stale = $cardData | Where-Object { (Get-ObjProp $_ 'alias') -eq $alias } | Select-Object -First 1
    if ($stale) {
        $staleToken = Get-ObjProp $stale 'token'
        if ($staleToken) {
            Remove-StaleSaltoAliasToken -Token $staleToken -UserId $UserId -Cfg $Cfg
        }
    }

    $csvPath = Join-Path $env:TEMP "unifi-nfc-$SaltoUserId.csv"
    Write-JsonNoBom $csvPath "$TagId,$alias`n"
    $authCfg = New-CurlAuthConfigFile -Token $Cfg.Token
    try {
        $importRaw = & curl.exe -sk -m 120 -K $authCfg -X POST "$api/credentials/nfc_cards/import" -F "file=@$csvPath" 2>&1
    } finally {
        Remove-Item -LiteralPath $authCfg -ErrorAction SilentlyContinue
    }
    $importText = ($importRaw | Out-String).Trim()
    if (-not $importText) { throw "NFC import returned empty response for ${TagId}" }
    $import = $importText | ConvertFrom-Json
    $importCode = Get-ObjProp $import 'code'
    if ($importCode -ne 'SUCCESS') {
        $importMsg = Get-ObjProp $import 'msg'
        throw "NFC import failed for ${TagId}: $importCode $importMsg"
    }
    $importItem = Get-FirstArrayItem (Get-ObjProp $import 'data')
    $token = Get-ObjProp $importItem 'token'
    if (-not $token) {
        throw "NFC import succeeded but no token returned for ${TagId}"
    }
    return $token
}

function Get-UserAssignedNfcTokens {
    param($UserDetail)

    $tokens = @()
    $cards = Get-ObjProp $UserDetail 'nfc_cards'
    if (-not $cards) { return ,@() }
    foreach ($card in @($cards)) {
        $token = Get-ObjProp $card 'token'
        if ($token) { $tokens += [string]$token }
    }
    return ,@($tokens | Select-Object -Unique)
}

function Get-PlanItemTagId {
    param([pscustomobject]$Item)

    foreach ($val in @($Item.UniFiTag, $Item.SaltoTag, (Get-ObjProp $Item.Desired 'tag_id'))) {
        $text = Normalize-SqlCmdValue ([string]$val)
        if (-not [string]::IsNullOrWhiteSpace($text) -and $text -ne '-') {
            return $text.ToUpper()
        }
    }
    return $null
}

function Find-NfcInventoryTokensForSaltoUser {
    param(
        [int]$SaltoUserId,
        [string]$TagId,
        $Cfg
    )

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $alias = "salto-$SaltoUserId"
    $tokens = @()
    $resp = Invoke-UniFiJson -Method GET -Url "$api/credentials/nfc_cards/tokens?page_size=500" -Token $Cfg.Token
    $cardData = Get-ObjProp $resp 'data'
    if (-not $cardData) { return ,@() }

    foreach ($card in @($cardData)) {
        $token = Get-ObjProp $card 'token'
        if (-not $token) { continue }
        $aliasVal = Get-ObjProp $card 'alias'
        $nfcId = Get-ObjProp $card 'nfc_id'
        $note = Get-ObjProp $card 'note'
        $match = ($aliasVal -eq $alias)
        if ($TagId -and $nfcId -and ([string]$nfcId).ToUpper() -eq $TagId) { $match = $true }
        if ($TagId -and $note -and ([string]$note).ToUpper() -eq $TagId) { $match = $true }
        if ($match) { $tokens += [string]$token }
    }

    return ,@($tokens | Select-Object -Unique)
}

function Test-NfcTokenExists {
    param(
        [string]$Token,
        $Cfg
    )

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    try {
        $resp = Invoke-UniFiJson -Method GET -Url "$api/credentials/nfc_cards/tokens/$Token" -Token $Cfg.Token
        return ($null -ne (Get-ObjProp $resp 'data'))
    } catch {
        return $false
    }
}

function Get-NfcTokensForSaltoUserCleanup {
    param(
        [pscustomobject]$Item,
        $Cfg
    )

    $tagId = Get-PlanItemTagId $Item
    $tokens = @()
    $tokens += Get-UserAssignedNfcTokens -UserDetail $Item.UniFiUser
    $tokens += Find-NfcInventoryTokensForSaltoUser -SaltoUserId $Item.id_user -TagId $tagId -Cfg $Cfg
    return ,@($tokens | Select-Object -Unique)
}

function Unassign-NfcTokensFromUser {
    param(
        [string]$UserId,
        [array]$Tokens,
        $Cfg,
        [string]$UserLabel
    )

    if (-not $UserId -or $Tokens.Count -eq 0) { return }

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    foreach ($token in $Tokens) {
        $path = Join-Path $env:TEMP "unifi-nfc-unassign-$token.json"
        $body = (@{ token = $token } | ConvertTo-Json -Compress)
        Write-JsonNoBom $path $body
        try {
            Invoke-UniFiJson -Method PUT -Url "$api/users/$UserId/nfc_cards/delete" -BodyPath $path -Token $Cfg.Token | Out-Null
            Write-Host "[NFC UNASSIGN] Removed token from $userLabel" -ForegroundColor Cyan
        } catch {
            Write-Host "[NFC UNASSIGN] Skipped token on $userLabel ($($_.Exception.Message))" -ForegroundColor DarkYellow
        }
    }
}

function Remove-NfcTokensFromInventory {
    param(
        [array]$Tokens,
        [int]$SaltoUserId,
        [string]$TagId,
        $Cfg,
        [string]$UserLabel
    )

    if (-not (Get-ConfigBool $Cfg 'DeleteOrphanNfcTokens' $true)) { return }

    $tokens = @($Tokens | Select-Object -Unique)
    $tokens += Find-NfcInventoryTokensForSaltoUser -SaltoUserId $SaltoUserId -TagId $TagId -Cfg $Cfg
    $tokens = @($tokens | Select-Object -Unique)

    if ($tokens.Count -eq 0) {
        Write-Host "[NFC] No tokens to delete from inventory for $userLabel" -ForegroundColor DarkGray
        return
    }

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    foreach ($token in $tokens) {
        if (-not (Test-NfcTokenExists -Token $token -Cfg $Cfg)) { continue }
        try {
            Invoke-UniFiJson -Method DELETE -Url "$api/credentials/nfc_cards/tokens/$token" -Token $Cfg.Token | Out-Null
        } catch {
            Write-Host "[NFC DELETE FAIL] Could not delete token for ${userLabel}: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
        if (Test-NfcTokenExists -Token $token -Cfg $Cfg) {
            Write-Host "[NFC DELETE FAIL] Token still in UniFi inventory after delete for $userLabel" -ForegroundColor Red
        } else {
            Write-Host "[NFC DELETE] Removed token from inventory for $userLabel" -ForegroundColor Magenta
        }
    }
}

function Assign-NfcToUser {
    param(
        [string]$UserId,
        [string]$TagId,
        [int]$SaltoUserId,
        $Cfg
    )
    if (-not $UserId) { throw 'Assign-NfcToUser requires UserId' }
    $token = Ensure-NfcToken -TagId $TagId -SaltoUserId $SaltoUserId -Cfg $Cfg -UserId $UserId
    if (-not $token) { throw "No NFC token resolved for tag ${TagId}" }

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $path = Join-Path $env:TEMP "unifi-nfc-assign-$SaltoUserId.json"
    $body = (@{ token = $token; force_add = $true } | ConvertTo-Json -Compress)
    Write-JsonNoBom $path $body
    Invoke-UniFiJson -Method PUT -Url "$api/users/$UserId/nfc_cards" -BodyPath $path -Token $Cfg.Token | Out-Null
}

function Invoke-SyncAction {
    param(
        [pscustomobject]$Item,
        $Cfg
    )

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $desired = $Item.Desired
    $userLabel = Format-SyncUserLabel $desired
    $userId = if ($Item.UniFiUser) { Get-ObjProp $Item.UniFiUser 'id' } else { $null }
    $actions = $Item.Action -split ',\s*'

    foreach ($action in $actions) {
        switch ($action) {
            'CREATE_USER' {
                $path = Join-Path $env:TEMP "unifi-create-$($desired.id_user).json"
                $body = (@{
                    first_name      = $desired.first_name
                    last_name       = $desired.last_name
                    employee_number = $desired.employee_number
                    status          = 'ACTIVE'
                } | ConvertTo-Json -Compress)
                Write-JsonNoBom $path $body
                $created = Invoke-UniFiJson -Method POST -Url "$api/users" -BodyPath $path -Token $Cfg.Token
                $userId = Get-ObjProp (Get-ObjProp $created 'data') 'id'
                Write-Host "[CREATE] $userLabel -> UniFi $userId" -ForegroundColor Green
            }
            'UPDATE_USER' {
                if (-not $userId) { throw "UPDATE_USER without UniFi user id for $userLabel" }
                $path = Join-Path $env:TEMP "unifi-update-$($desired.id_user).json"
                $body = (@{
                    first_name      = $desired.first_name
                    last_name       = $desired.last_name
                    employee_number = $desired.employee_number
                    status          = 'ACTIVE'
                } | ConvertTo-Json -Compress)
                Write-JsonNoBom $path $body
                Invoke-UniFiJson -Method PUT -Url "$api/users/$userId" -BodyPath $path -Token $Cfg.Token | Out-Null
                Write-Host "[UPDATE] $userLabel name/employee_number" -ForegroundColor Yellow
            }
            'ACTIVATE_USER' {
                if (-not $userId) { throw "ACTIVATE_USER without UniFi user id for $userLabel" }
                $path = Join-Path $env:TEMP "unifi-activate-$($desired.id_user).json"
                $body = (@{
                    first_name      = $desired.first_name
                    last_name       = $desired.last_name
                    employee_number = $desired.employee_number
                    status          = 'ACTIVE'
                } | ConvertTo-Json -Compress)
                Write-JsonNoBom $path $body
                Invoke-UniFiJson -Method PUT -Url "$api/users/$userId" -BodyPath $path -Token $Cfg.Token | Out-Null
                Write-Host "[ACTIVATE] $userLabel" -ForegroundColor Yellow
            }
            'ADD_GROUP' {
                if (-not $userId) { throw "ADD_GROUP without UniFi user id for $userLabel" }
                $targetGroupId = $Item.GroupId
                if (-not $targetGroupId) { throw "ADD_GROUP without group id for $userLabel" }
                $path = Join-Path $env:TEMP "unifi-group-$($desired.id_user).json"
                Write-JsonNoBom $path "[`"$userId`"]"
                Invoke-UniFiJson -Method POST -Url "$api/user_groups/$targetGroupId/users" -BodyPath $path -Token $Cfg.Token | Out-Null
                Write-Host "[GROUP] Added $userLabel to $($Cfg.UserGroupName)" -ForegroundColor Cyan
            }
            'ASSIGN_NFC' {
                if (-not $userId) { throw "ASSIGN_NFC without UniFi user id for $userLabel" }
                Assign-NfcToUser -UserId $userId -TagId $desired.tag_id -SaltoUserId $desired.id_user -Cfg $Cfg
                Write-Host "[NFC] Assigned tag $($desired.tag_id) to $userLabel" -ForegroundColor Cyan
            }
            'UPDATE_NFC' {
                if (-not $userId) { throw "UPDATE_NFC without UniFi user id for $userLabel" }
                Assign-NfcToUser -UserId $userId -TagId $desired.tag_id -SaltoUserId $desired.id_user -Cfg $Cfg
                Write-Host "[NFC] Updated tag $($desired.tag_id) for $userLabel" -ForegroundColor Yellow
            }
            'REFRESH' {
                Write-Host "[REFRESH] $userLabel unchanged fields re-checked" -ForegroundColor DarkGray
            }
            'NONE' {
                Write-Host "[OK] $userLabel already in sync" -ForegroundColor DarkGreen
            }
            'SKIP_NON_BASIC' {
                Write-Host "[SKIP] ${userLabel}: $($Item.Details)" -ForegroundColor Red
            }
            'DEACTIVATE_USER' {
                if (-not $userId) { throw "DEACTIVATE_USER without UniFi user id for $userLabel" }
                $path = Join-Path $env:TEMP "unifi-deactivate-$($Item.id_user).json"
                $body = (@{
                    first_name      = (Get-ObjProp $Item.UniFiUser 'first_name')
                    last_name       = (Get-ObjProp $Item.UniFiUser 'last_name')
                    employee_number = (Get-ObjProp $Item.UniFiUser 'employee_number')
                    status          = 'DEACTIVATED'
                } | ConvertTo-Json -Compress)
                Write-JsonNoBom $path $body
                Invoke-UniFiJson -Method PUT -Url "$api/users/$userId" -BodyPath $path -Token $Cfg.Token | Out-Null
                Write-Host "[DEACTIVATE] $userLabel ($($Item.Details))" -ForegroundColor Magenta
            }
            'DELETE_USER' {
                if (-not $userId) { throw "DELETE_USER without UniFi user id for $userLabel" }
                $tagId = Get-PlanItemTagId $Item
                $nfcTokens = Get-NfcTokensForSaltoUserCleanup -Item $Item -Cfg $Cfg
                Unassign-NfcTokensFromUser -UserId $userId -Tokens $nfcTokens -Cfg $Cfg -UserLabel $userLabel
                if ((Get-ObjProp $Item.UniFiUser 'status') -eq 'ACTIVE') {
                    $path = Join-Path $env:TEMP "unifi-deactivate-before-delete-$($Item.id_user).json"
                    $body = (@{
                        first_name      = (Get-ObjProp $Item.UniFiUser 'first_name')
                        last_name       = (Get-ObjProp $Item.UniFiUser 'last_name')
                        employee_number = (Get-ObjProp $Item.UniFiUser 'employee_number')
                        status          = 'DEACTIVATED'
                    } | ConvertTo-Json -Compress)
                    Write-JsonNoBom $path $body
                    Invoke-UniFiJson -Method PUT -Url "$api/users/$userId" -BodyPath $path -Token $Cfg.Token | Out-Null
                }
                Invoke-UniFiJson -Method DELETE -Url "$api/users/$userId" -Token $Cfg.Token | Out-Null
                Write-Host "[DELETE] $userLabel ($($Item.Details))" -ForegroundColor Magenta
                Remove-NfcTokensFromInventory -Tokens $nfcTokens -SaltoUserId $Item.id_user -TagId $tagId -Cfg $Cfg -UserLabel $userLabel
            }
            default {
                if ($action) {
                    Write-Host "[WARN] Unknown action '$action' for $userLabel" -ForegroundColor Red
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Self-update (GitHub)
# ---------------------------------------------------------------------------

function Compare-ScriptVersion {
    param(
        [string]$Local,
        [string]$Remote
    )

    if ([string]::IsNullOrWhiteSpace($Remote)) { return 1 }
    if ([string]::IsNullOrWhiteSpace($Local)) { return -1 }

    try {
        $localVer = [version]($Local -replace '-.*$', '')
        $remoteVer = [version]($Remote -replace '-.*$', '')
        return $localVer.CompareTo($remoteVer)
    } catch {
        return [string]::Compare($Local, $Remote, [System.StringComparison]::OrdinalIgnoreCase)
    }
}

function Get-UpdateCheckCachePath($Cfg) {
    $logDir = Get-ObjProp $Cfg 'LogDir'
    if (-not $logDir) { $logDir = Get-DefaultLogDir }
    return (Join-Path $logDir 'last-update-check.txt')
}

function Test-UpdateCheckDue {
    param(
        $Cfg,
        [switch]$Force
    )

    if ($Force) { return $true }

    $intervalHours = Get-ObjProp $Cfg 'UpdateCheckIntervalHours'
    if ($null -eq $intervalHours) { $intervalHours = 24 }
    if ($intervalHours -le 0) { return $true }

    $cachePath = Get-UpdateCheckCachePath -Cfg $Cfg
    if (-not (Test-Path -LiteralPath $cachePath)) { return $true }

    try {
        $lastCheck = [datetime]::Parse((Get-Content -LiteralPath $cachePath -Raw).Trim())
        return ((Get-Date) - $lastCheck).TotalHours -ge [double]$intervalHours
    } catch {
        return $true
    }
}

function Set-UpdateCheckTimestamp {
    param($Cfg)

    $logDir = Get-ObjProp $Cfg 'LogDir'
    if (-not $logDir) { $logDir = Get-DefaultLogDir }
    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $cachePath = Get-UpdateCheckCachePath -Cfg $Cfg
    Set-Content -LiteralPath $cachePath -Value (Get-Date).ToString('o') -Encoding UTF8
}

function Get-FileSha256Hex {
    param([Parameter(Mandatory)][string]$Path)

    $hash = Get-FileHash -LiteralPath $Path -Algorithm SHA256
    return $hash.Hash.ToLowerInvariant()
}

function Invoke-GitHubRepoDownload {
    param(
        [Parameter(Mandatory)][string]$Owner,
        [Parameter(Mandatory)][string]$Repo,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Ref,
        [Parameter(Mandatory)][string]$OutPath
    )

    $url = "https://api.github.com/repos/$Owner/$Repo/contents/$Path" + "?ref=$Ref"
    $curlArgs = @(
        '-skL', '-m', '120',
        '-H', 'Accept: application/vnd.github.raw+json',
        '-H', 'User-Agent: salto-unifi-sync',
        '-o', $OutPath,
        $url
    )

    $raw = & curl.exe @curlArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "curl failed (exit $LASTEXITCODE): $(($raw | Out-String).Trim())"
    }
    if (-not (Test-Path -LiteralPath $OutPath)) {
        throw "Download missing: $Path ($Ref)"
    }
    if ((Get-Item -LiteralPath $OutPath).Length -eq 0) {
        throw "Download empty: $Path ($Ref)"
    }
}

function Get-RemoteVersionManifest {
    param(
        $Cfg,
        [string]$Channel
    )

    $owner = [string](Get-ObjProp $Cfg 'UpdateRepoOwner')
    $repo = [string](Get-ObjProp $Cfg 'UpdateRepoName')
    if (-not $owner -or -not $repo) {
        throw 'UpdateRepoOwner and UpdateRepoName must be set when AutoUpdate is enabled.'
    }

    $tempFile = Join-Path $env:TEMP ("salto-unifi-version-{0}.json" -f ([guid]::NewGuid().ToString('N')))

    try {
        Invoke-GitHubRepoDownload -Owner $owner -Repo $repo -Path 'version.json' -Ref $Channel `
            -OutPath $tempFile
        $raw = Get-Content -LiteralPath $tempFile -Raw
        return ($raw | ConvertFrom-Json)
    } finally {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
}

function Get-RemoteScriptSha256($Manifest) {
    $fileEntry = Get-ObjProp $Manifest 'files'
    if ($fileEntry) {
        $scriptEntry = Get-ObjProp $fileEntry 'Sync-Salto-Unifi.ps1'
        if ($scriptEntry) {
            $hash = Get-ObjProp $scriptEntry 'sha256'
            if ($hash) { return [string]$hash.ToLowerInvariant() }
        }
    }
    $legacy = Get-ObjProp $Manifest 'sha256'
    if ($legacy) { return [string]$legacy.ToLowerInvariant() }
    return $null
}

function Exit-AfterScriptUpdate {
    param(
        [string]$ScriptPath,
        [string]$RemoteVersion,
        [string]$LocalVersion,
        [string]$Mode,
        [string]$Filter,
        [string]$ConfigPath,
        [string]$BackupPath
    )

    $scriptDir = [System.IO.Path]::GetDirectoryName($ScriptPath)
    $scriptName = [System.IO.Path]::GetFileName($ScriptPath)
    $backupName = [System.IO.Path]::GetFileName($BackupPath)

    Write-Host ''
    Write-Host '[UPDATE] New version installed successfully.' -ForegroundColor Green
    Write-Host "[UPDATE] Installed : v$RemoteVersion (file on disk)" -ForegroundColor Green
    Write-Host "[UPDATE] This run : v$LocalVersion (still in memory - sync not started)" -ForegroundColor Yellow
    Write-Host "[UPDATE] Backup   : $backupName" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '[UPDATE] Re-run the script to sync with the new version:' -ForegroundColor Cyan
    Write-Host "  cd '$scriptDir'" -ForegroundColor White
    Write-Host "  .\$scriptName -Mode $Mode -Filter '$Filter' -ConfigPath '$ConfigPath'" -ForegroundColor White
    Write-Host ''
    exit 0
}

function Invoke-ScriptSelfUpdate {
    param(
        $Cfg,
        [switch]$ForceUpdateCheck,
        [string]$Mode,
        [string]$Filter,
        [string]$ConfigPath
    )

    if (-not (Get-ConfigBool $Cfg 'AutoUpdate' $false)) { return $false }
    if (-not (Test-UpdateCheckDue -Cfg $Cfg -Force:$ForceUpdateCheck)) {
        Write-Host '[UPDATE] Skipped (checked recently; use -ForceUpdateCheck to override).' -ForegroundColor DarkGray
        return $false
    }

    $channel = [string](Get-ObjProp $Cfg 'UpdateChannel')
    if (-not $channel) { $channel = 'main' }

    Write-Host "[UPDATE] Checking GitHub channel '$channel' for newer script..." -ForegroundColor Cyan

    try {
        $manifest = Get-RemoteVersionManifest -Cfg $Cfg -Channel $channel
        $remoteVersion = [string](Get-ObjProp $manifest 'version')
        if (-not $remoteVersion) {
            throw 'Remote version.json has no version field.'
        }

        if ((Compare-ScriptVersion -Local $ScriptVersion -Remote $remoteVersion) -ge 0) {
            Write-Host "[UPDATE] Already on v$ScriptVersion (remote v$remoteVersion)." -ForegroundColor DarkGray
            Set-UpdateCheckTimestamp -Cfg $Cfg
            return $false
        }

        $expectedHash = Get-RemoteScriptSha256 -Manifest $manifest
        if (-not $expectedHash) {
            throw 'Remote version.json has no SHA256 for Sync-Salto-Unifi.ps1.'
        }

        $owner = [string](Get-ObjProp $Cfg 'UpdateRepoOwner')
        $repo = [string](Get-ObjProp $Cfg 'UpdateRepoName')
        $scriptPath = Join-Path $PSScriptRoot 'Sync-Salto-Unifi.ps1'
        $newPath = "$scriptPath.new"
        $bakPath = "$scriptPath.bak"

        Write-Host "[UPDATE] Downloading v$remoteVersion..." -ForegroundColor Cyan
        Invoke-GitHubRepoDownload -Owner $owner -Repo $repo -Path 'Sync-Salto-Unifi.ps1' -Ref $channel `
            -OutPath $newPath

        $actualHash = Get-FileSha256Hex -Path $newPath
        if ($actualHash -ne $expectedHash) {
            Remove-Item -LiteralPath $newPath -Force -ErrorAction SilentlyContinue
            throw "SHA256 mismatch for downloaded script (expected $expectedHash, got $actualHash)."
        }

        Copy-Item -LiteralPath $scriptPath -Destination $bakPath -Force
        Move-Item -LiteralPath $newPath -Destination $scriptPath -Force
        Set-UpdateCheckTimestamp -Cfg $Cfg
        try {
            Exit-AfterScriptUpdate -ScriptPath $scriptPath -RemoteVersion $remoteVersion `
                -LocalVersion $ScriptVersion -Mode $Mode -Filter $Filter -ConfigPath $ConfigPath `
                -BackupPath $bakPath
        } catch {
            Write-Host ''
            Write-Host '[UPDATE] New version installed on disk; this process still runs the old script in memory.' -ForegroundColor Green
            Write-Host "[UPDATE] Installed: v$remoteVersion | This run: v$ScriptVersion" -ForegroundColor Yellow
            Write-Host '[UPDATE] Re-run the script to continue:' -ForegroundColor Cyan
            $dir = [System.IO.Path]::GetDirectoryName($scriptPath)
            $name = [System.IO.Path]::GetFileName($scriptPath)
            Write-Host "  cd '$dir'" -ForegroundColor White
            Write-Host "  .\$name -Mode $Mode -Filter '$Filter' -ConfigPath '$ConfigPath'" -ForegroundColor White
            Write-Host ''
            exit 0
        }
    } catch {
        Write-Host "[UPDATE FAIL] $($_.Exception.Message)" -ForegroundColor Red
        Write-Host '[UPDATE] Continuing with current script version.' -ForegroundColor DarkYellow
        return $false
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$runId = (Get-Date).ToString('yyyyMMdd_HHmmss')
$runSummary = ''
$runStatus = 'SUCCESS'
$runError = $null
$script:RunBoundaryStarted = $false

try {
    $cfg = Read-Config -Path $ConfigPath

    if (-not $SkipUpdate) {
        Invoke-ScriptSelfUpdate -Cfg $cfg -ForceUpdateCheck:$ForceUpdateCheck `
            -Mode $Mode -Filter $Filter -ConfigPath $ConfigPath | Out-Null
    }

    # "SqlRunAs": run this script again as that account (a private LocalDB instance
    # can only be opened by the account that owns it) and stream its output here.
    $runAs = ([string](Get-ObjProp $cfg 'SqlRunAs')).Trim()
    if ($runAs -and -not $NoDelegate -and -not (Test-RunningAsAccount -Account $runAs)) {
        # This (elevated) session writes the log of the whole run: the account that owns
        # the LocalDB instance usually cannot write to the log folder.
        $parentLog = Start-RunLogging -Cfg $cfg -RunId $runId
        $childParams = @{ ConfigPath = $ConfigPath; Mode = $Mode; Filter = $Filter; NoDelegate = $true; SkipUpdate = $true; NoTranscript = $true }
        if ($parentLog) { $childParams['ParentLogPath'] = [string]$parentLog }
        if ($Force) { $childParams['Force'] = $true }
        if ($ShowAll) { $childParams['ShowAll'] = $true }
        $delegatedExit = [int](@(Invoke-DelegatedRun -Account $runAs -ChildParams $childParams -ScriptFile $PSCommandPath) | Select-Object -Last 1)
        Stop-RunLogging
        exit $delegatedExit
    }

    if ($NoTranscript) {
        $logPath = $ParentLogPath
        if ($logPath) { $script:RunLogPath = $logPath }
    } else {
        $logPath = Start-RunLogging -Cfg $cfg -RunId $runId
    }
    $updateChannel = Get-UpdateChannelLabel -Cfg $cfg
    Write-RunBoundary -Phase BEGIN -Mode $Mode -Filter $Filter -UpdateChannel $updateChannel -LogPath $logPath -ConfigFile $ConfigPath
    $script:RunBoundaryStarted = $true

    $idFilter = Parse-IdFilter -FilterText $Filter

    Test-SqlServerPreflight -Cfg $cfg

    Write-Host 'Loading Salto users (sync scope)...'
    $saltoUsers = @(Get-SaltoUsers -Cfg $cfg -IdFilter $idFilter)
    Write-Host "Salto users in sync scope: $($saltoUsers.Count)"

    if (Get-ConfigBool $cfg 'DeactivateWhenIneligible' $true) {
        Write-Host 'Loading Salto users (lookup for deactivation)...'
        $saltoLookup = @(Get-SaltoUsers -Cfg $cfg -IdFilter $idFilter -LookupMode)
        Write-Host "Salto users in lookup scope: $($saltoLookup.Count)"
    } else {
        $saltoLookup = @()
    }

    Write-Host 'Loading UniFi users...'
    $unifiUsers = @(Get-AllUniFiUsers -Cfg $cfg)
    Write-Host ('UniFi users loaded: {0}' -f $unifiUsers.Count)

    $plan = Build-SyncPlan `
        -SaltoUsers $saltoUsers `
        -SaltoLookup $saltoLookup `
        -UniFiUsers $unifiUsers `
        -Cfg $cfg `
        -SyncMode $Mode `
        -IdFilter $idFilter

    Write-SyncPlan -Plan $plan -ShowAll:$ShowAll
    if ($idFilter.Type -eq 'All' -and ($Mode -eq 'ShowDiff' -or $ShowAll)) {
        Write-UniFiReconciliation -Plan $plan -UniFiUsers $unifiUsers -SaltoCount $saltoUsers.Count
    }

    $toApply = @($plan | Where-Object { $_.Action -notin @('NONE', 'SKIP_NON_BASIC') })

    # Safety check: a Salto/SQL data problem (wrong filter, bad connection, typo in
    # SaltoUserType, ...) can make every managed user look "ineligible" at once. If a
    # large fraction of the previously-managed population would be deactivated or
    # deleted in one run, abort rather than silently wiping access for everyone.
    # Only applies to full-population runs (-Filter ALL); scoped test runs are exempt.
    $ineligibleActions = @($plan | Where-Object { $_.Action -match 'DEACTIVATE_USER|DELETE_USER' })
    if ($idFilter.Type -eq 'All' -and $ineligibleActions.Count -gt 0) {
        $baseline = $saltoUsers.Count + $ineligibleActions.Count
        $maxPct = [double](Get-ObjProp $cfg 'MaxDeactivationPercent')
        if ($maxPct -le 0) { $maxPct = 10 }
        $impactPct = if ($baseline -gt 0) { [math]::Round(100.0 * $ineligibleActions.Count / $baseline, 1) } else { 0 }
        $impactColor = if ($impactPct -gt $maxPct) { 'Red' } else { 'DarkGray' }
        Write-Host ''
        Write-Host ("Deactivate/delete impact: {0} of {1} managed user(s) ({2}%, threshold {3}%)" -f $ineligibleActions.Count, $baseline, $impactPct, $maxPct) -ForegroundColor $impactColor
        if ($impactPct -gt $maxPct -and $Mode -ne 'ShowDiff' -and -not $Force) {
            throw "Safety check failed: $($ineligibleActions.Count) of $baseline managed user(s) ($impactPct%) would be deactivated or deleted, exceeding the $maxPct% MaxDeactivationPercent threshold. This usually points to a Salto/SQL data problem rather than a real change. Review with -Mode ShowDiff, and re-run with -Force only after confirming the plan is correct."
        }
    }

    if ($Mode -eq 'ShowDiff') {
        $runSummary = "ShowDiff complete. $($toApply.Count) user(s) would be changed."
        Write-Host ''
        Write-Host $runSummary -ForegroundColor Cyan
        return
    }

    if ($toApply.Count -eq 0) {
        $runSummary = 'Nothing to sync.'
        Write-Host $runSummary -ForegroundColor Green
        return
    }

    Write-Host ''
    Write-Host "Applying $($toApply.Count) change set(s)..." -ForegroundColor White
    foreach ($item in ($toApply | Sort-Object id_user)) {
        Invoke-SyncAction -Item $item -Cfg $cfg
    }

    $runSummary = "Applied $($toApply.Count) change set(s)."
    Write-Host ''
    Write-Host '=== Done ===' -ForegroundColor Green
}
catch {
    if (-not $script:RunBoundaryStarted) {
        if ($NoTranscript) {
            $logPath = $ParentLogPath
            if ($logPath) { $script:RunLogPath = $logPath }
        } else {
            $logPath = Start-RunLogging -Cfg $null -RunId $runId
        }
        Write-RunBoundary -Phase BEGIN -Mode $Mode -Filter $Filter -UpdateChannel 'n/a' -LogPath $logPath -ConfigFile $ConfigPath
        $script:RunBoundaryStarted = $true
    }
    $runStatus = 'FAILED'
    $runError = $_
    $runSummary = $_.Exception.Message
    Write-Host ''
    Write-Host "FATAL: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
finally {
    if ($script:RunBoundaryStarted) {
        Write-RunBoundary -Phase END -Status $runStatus -Summary $runSummary -ErrorRecord $runError
    }
    Stop-RunLogging
}
