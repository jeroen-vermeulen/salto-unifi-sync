<#
.SYNOPSIS
    One-way sync from Salto ProAccess Space (SALTO_SPACE) to UniFi Access.

.DESCRIPTION
    Salto is the source of truth. Only basic local UniFi users managed by this
    script (matched via employee_number = Salto id_user) are created or updated.
    Non-basic UniFi accounts are never modified.
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
    Skip the GitHub auto-update check (used internally after a successful update restart).

.PARAMETER ForceUpdateCheck
    Check GitHub for updates even if UpdateCheckIntervalHours has not elapsed.

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

    [switch]$ForceUpdateCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent -LiteralPath $MyInvocation.MyCommand.Path
}
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot 'unifi-sync-config.json'
}
$ScriptVersion = '1.3.3'

$script:RunLogPath = $null
$script:TranscriptActive = $false

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

function Invoke-SaltoSqlCmd {
    param(
        [string]$SqlServer,
        [string]$Database,
        [string]$InputFile
    )

    $previousOutputEncoding = [Console]::OutputEncoding
    $previousPsOutputEncoding = $OutputEncoding
    try {
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        [Console]::OutputEncoding = $utf8
        $OutputEncoding = $utf8
        $raw = & sqlcmd.exe -S $SqlServer -E -d $Database -W -s '|' -i $InputFile -f i:65001,o:65001 2>&1
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
    Set-ConfigDefault $cfg 'AutoUpdate' $false
    Set-ConfigDefault $cfg 'UpdateChannel' 'main'
    Set-ConfigDefault $cfg 'UpdateRepoOwner' 'jeroen-vermeulen'
    Set-ConfigDefault $cfg 'UpdateRepoName' 'salto-unifi-sync'
    Set-ConfigDefault $cfg 'UpdateCheckIntervalHours' 24
    $ghTokenFile = Get-ObjProp $cfg 'UpdateGitHubTokenFile'
    if ($ghTokenFile -and (Test-Path -LiteralPath $ghTokenFile)) {
        $ghToken = (Get-Content -LiteralPath $ghTokenFile -Raw).Trim()
        if ($ghToken) {
            $cfg | Add-Member -NotePropertyName UpdateGitHubToken -NotePropertyValue $ghToken -Force
        }
    }
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

    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    $logPath = Join-Path $logDir "salto-unifi-sync_$RunId.log"
    Start-Transcript -Path $logPath -Force | Out-Null
    $script:RunLogPath = $logPath
    $script:TranscriptActive = $true
    return $logPath
}

function Stop-RunLogging {
    if ($script:TranscriptActive) {
        Stop-Transcript | Out-Null
        $script:TranscriptActive = $false
    }
}

function Write-RunBoundary {
    param(
        [ValidateSet('BEGIN', 'END')]
        [string]$Phase,

        [string]$Mode = '',
        [string]$Filter = '',
        [string]$LogPath = '',
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
        Write-Host "Mode      : $Mode"
        Write-Host "Filter    : $Filter"
        Write-Host "User      : $env:USERDOMAIN\$env:USERNAME"
        Write-Host "Computer  : $env:COMPUTERNAME"
        if ($LogPath) {
            Write-Host "Log file  : $LogPath"
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

function Get-IneligibleAction {
    param([string]$Reason)
    switch ($Reason) {
        'NO_SALTO_TAG' { return 'DELETE_USER' }
        'SALTO_DELETED' { return 'DELETE_USER' }
        'NOT_IN_SALTO' { return 'DELETE_USER' }
        default { return 'DEACTIVATE_USER' }
    }
}

function Invoke-UniFiJson {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Url,
        [string]$BodyPath = $null,
        [string]$Token
    )
    $respFile = Join-Path $env:TEMP ("unifi-resp-{0}.json" -f [guid]::NewGuid().ToString('N'))
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    try {
        $args = @(
            '-sk', '-m', '120',
            '-X', $Method,
            $Url,
            '-H', "Authorization: Bearer $Token",
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

    $expectedAlias = "salto-$($Desired.id_user)"
    $cards = Get-ObjProp $UniFiUserDetail 'nfc_cards'
    if (-not $cards) { return $false }

    foreach ($card in @($cards)) {
        $token = Get-ObjProp $card 'token'
        if (-not $token) { continue }
        $alias = $null
        if ($NfcTokenMap -and $NfcTokenMap.ContainsKey("token_alias:$token")) {
            $alias = $NfcTokenMap["token_alias:$token"]
        }
        if ($alias -eq $expectedAlias) { return $true }
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

    $api = "$($Cfg.UnifiHost)/api/v1/developer"
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

        $detailResp = Invoke-UniFiJson -Method GET -Url "$api/users/$userId" -Token $Cfg.Token
        $detail = Get-ObjProp $detailResp 'data'
        if (-not $detail) { continue }

        $empNo = Get-ObjProp $detail 'employee_number'
        if ($empNo -match '^\d+$') {
            $index.ByEmployee[[string]$empNo] = $detail
        }

        $firstName = Get-ObjProp $detail 'first_name'
        $lastName = Get-ObjProp $detail 'last_name'
        if ($firstName -and $lastName) {
            $index.ByName["$firstName|$lastName"] = $detail
        }

        $tag = Get-UserNfcUid $detail -NfcTokenMap $nfcTokenMap
        if ($tag) {
            $index.ByTag[$tag] = $detail
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

function Test-UserInGroup {
    param(
        [string]$GroupId,
        [string]$UserId,
        $Cfg
    )
    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $resp = Invoke-UniFiJson -Method GET -Url "$api/user_groups/$GroupId/users" -Token $Cfg.Token
    $data = Get-ObjProp $resp 'data'
    if (-not $data) { return $false }
    return @($data | Where-Object { (Get-ObjProp $_ 'id') -eq $UserId }).Count -gt 0
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
        [string]$Api
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

    $plan = @()
    foreach ($salto in $SaltoUsers) {
        $desired = Get-DesiredState $salto
        $actions = @()
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
                $plan += [pscustomobject]@{
                    id_user   = $desired.id_user
                    Name      = (Format-SyncUserLabel $desired)
                    SaltoTag  = $desired.tag_id
                    UniFiTag  = (Get-UserNfcDisplayTag $uniDetail $desired $nfcTokenMap)
                    Match     = $matchBy
                    Status    = $desired.status_text
                    Action    = 'SKIP_NON_BASIC'
                    Details   = $basic.Reason
                    Desired   = $desired
                    UniFiUser = $uniUser
                }
                continue
            }

            $uniSummary = "$(Get-ObjProp $uniDetail 'first_name') $(Get-ObjProp $uniDetail 'last_name')"
            $currentTag = Get-UserNfcDisplayTag $uniDetail $desired $nfcTokenMap
            $inGroup = Test-UserInGroup -GroupId $groupId -UserId $uniUserId -Cfg $Cfg

            $nameDiff = Test-NameDiffers `
                -UniFirst (Get-ObjProp $uniDetail 'first_name') `
                -UniLast (Get-ObjProp $uniDetail 'last_name') `
                -DesiredFirst $desired.first_name `
                -DesiredLast $desired.last_name
            $empDiff = ((Get-ObjProp $uniDetail 'employee_number') -ne $desired.employee_number)
            $tagDiff = -not (Test-NfcInSync $uniDetail $desired $nfcTokenMap)
            $groupDiff = -not $inGroup
            $statusDiff = ((Get-ObjProp $uniDetail 'status') -ne 'ACTIVE')

            if ($SyncMode -eq 'FullSync') {
                if ($nameDiff -or $empDiff) { $actions += 'UPDATE_USER' }
                if ($tagDiff) { $actions += 'UPDATE_NFC' }
                if ($groupDiff) { $actions += 'ADD_GROUP' }
                if ($statusDiff) { $actions += 'ACTIVATE_USER' }
                if ($actions.Count -eq 0) { $actions += 'REFRESH' }
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
            Desired   = $desired
            UniFiUser = $uniUser
            GroupId   = $groupId
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
        -Api $api

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

function Ensure-NfcToken {
    param(
        [string]$TagId,
        [int]$SaltoUserId,
        $Cfg
    )
    $api = "$($Cfg.UnifiHost)/api/v1/developer"
    $alias = "salto-$SaltoUserId"
    $cards = Invoke-UniFiJson -Method GET -Url "$api/credentials/nfc_cards/tokens?page_size=500" -Token $Cfg.Token
    $cardData = Get-ObjProp $cards 'data'
    $existing = $cardData | Where-Object {
        $aliasVal = Get-ObjProp $_ 'alias'
        $nfcId = Get-ObjProp $_ 'nfc_id'
        $note = Get-ObjProp $_ 'note'
        ($aliasVal -eq $alias) -or ($nfcId -and ([string]$nfcId).ToUpper() -eq $TagId) -or ($note -and ([string]$note).ToUpper() -eq $TagId)
    } | Select-Object -First 1
    if ($existing) {
        $token = Get-ObjProp $existing 'token'
        if ($token) { return $token }
    }

    $csvPath = Join-Path $env:TEMP "unifi-nfc-$SaltoUserId.csv"
    Write-JsonNoBom $csvPath "$TagId,$alias`n"
    $importRaw = & curl.exe -sk -m 120 -X POST "$api/credentials/nfc_cards/import" -H "Authorization: Bearer $($Cfg.Token)" -F "file=@$csvPath" 2>&1
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
    $token = Ensure-NfcToken -TagId $TagId -SaltoUserId $SaltoUserId -Cfg $Cfg
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
        [Parameter(Mandatory)][string]$OutPath,
        [string]$GitHubToken
    )

    $url = "https://api.github.com/repos/$Owner/$Repo/contents/$Path" + "?ref=$Ref"
    $curlArgs = @(
        '-skL', '-m', '120',
        '-H', 'Accept: application/vnd.github.raw+json',
        '-H', 'User-Agent: salto-unifi-sync',
        '-o', $OutPath,
        $url
    )
    if ($GitHubToken) {
        $curlArgs = @(
            '-skL', '-m', '120',
            '-H', "Authorization: Bearer $GitHubToken",
            '-H', 'Accept: application/vnd.github.raw+json',
            '-H', 'User-Agent: salto-unifi-sync',
            '-o', $OutPath,
            $url
        )
    }

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
    $ghToken = Get-ObjProp $Cfg 'UpdateGitHubToken'

    try {
        Invoke-GitHubRepoDownload -Owner $owner -Repo $repo -Path 'version.json' -Ref $Channel `
            -OutPath $tempFile -GitHubToken $ghToken
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

function Restart-AfterScriptUpdate {
    param(
        [string]$ScriptPath,
        [string]$RemoteVersion,
        [string]$Mode,
        [string]$Filter,
        [string]$ConfigPath
    )

    Write-Host "[UPDATE] Restarting with v$RemoteVersion..." -ForegroundColor Green
    $scriptDir = Split-Path -Parent -LiteralPath $ScriptPath
    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-Mode', $Mode,
        '-Filter', $Filter,
        '-ConfigPath', $ConfigPath,
        '-SkipUpdate'
    )
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -WorkingDirectory $scriptDir | Out-Null
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
        $ghToken = Get-ObjProp $Cfg 'UpdateGitHubToken'

        Write-Host "[UPDATE] Downloading v$remoteVersion..." -ForegroundColor Cyan
        Invoke-GitHubRepoDownload -Owner $owner -Repo $repo -Path 'Sync-Salto-Unifi.ps1' -Ref $channel `
            -OutPath $newPath -GitHubToken $ghToken

        $actualHash = Get-FileSha256Hex -Path $newPath
        if ($actualHash -ne $expectedHash) {
            Remove-Item -LiteralPath $newPath -Force -ErrorAction SilentlyContinue
            throw "SHA256 mismatch for downloaded script (expected $expectedHash, got $actualHash)."
        }

        Copy-Item -LiteralPath $scriptPath -Destination $bakPath -Force
        Move-Item -LiteralPath $newPath -Destination $scriptPath -Force
        Write-Host "[UPDATE] Installed v$remoteVersion (backup: $(Split-Path -Leaf $bakPath))." -ForegroundColor Green
        Set-UpdateCheckTimestamp -Cfg $Cfg
        Restart-AfterScriptUpdate -ScriptPath $scriptPath -RemoteVersion $remoteVersion `
            -Mode $Mode -Filter $Filter -ConfigPath $ConfigPath
        return $true
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

    $logPath = Start-RunLogging -Cfg $cfg -RunId $runId
    Write-RunBoundary -Phase BEGIN -Mode $Mode -Filter $Filter -LogPath $logPath
    $script:RunBoundaryStarted = $true

    $idFilter = Parse-IdFilter -FilterText $Filter

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

    Write-Host ''
    Write-Host '=== Plan ===' -ForegroundColor White
    $plan | Sort-Object id_user | Format-Table id_user, Name, Status, SaltoTag, UniFiTag, Match, Action, Details -AutoSize

    $toApply = @($plan | Where-Object { $_.Action -notin @('NONE', 'SKIP_NON_BASIC') })
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
        $logPath = Start-RunLogging -Cfg $null -RunId $runId
        Write-RunBoundary -Phase BEGIN -Mode $Mode -Filter $Filter -LogPath $logPath
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
