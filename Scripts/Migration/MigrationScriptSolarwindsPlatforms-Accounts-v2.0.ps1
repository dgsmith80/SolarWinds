#Requires -Version 5.1
#Requires -Modules SwisPowerShell
<#
.SYNOPSIS
    Migrates user accounts (Windows users/groups and Orion accounts) with their rights, limitations,
    home/summary views and user settings from one SolarWinds Platform (Orion) instance to another via SWIS.

.DESCRIPTION
    Author:     Holger Mundt
    Email:      holger.mundt@loop1.com
    Version:    2.0
    History:    v1.0    Original "Orion - Account Migration" script (OrionSDK sample based)
                v2.0    Refactoring of v1.0:
                          - no hard-coded IPs / test filter (AccountID = 'brian sanner') / initial password "Welcome1"
                            (parameters, Get-Credential or -Trusted; -NameFilter; -InitialPassword)
                          - the two copy-pasted branches (Windows / Orion account) are one code path
                          - typos fixed that silently dropped settings ($AllowDisableActio, $AllowAlertManagment);
                            "Enabled" and "CanClearEvents" were read but never applied - now they are
                          - limitations: handled per slot (1-3) instead of nested ifs, NULL-safe SQL escaping
                          - an account whose limitations cannot be migrated is NOT created (dropping a limitation
                            would give the account MORE access than on the source)
                          - a half-configured account is removed again, so a re-run can retry it
                          - -WhatIf support, log file, summary, exit code

    Runs standalone. Accounts that already exist on the target (same AccountID) are skipped.
    Run it AFTER the platform migration (custom properties, nodes, groups): limitations and views
    are matched by name/title against what already exists on the target.

    - Account types: 1 = Orion account, 2 = Windows user, 3 = Windows group. Other types are skipped (warning).
    - Home page view and summary view are matched by title + view group name; if not found the default (ID 1) is used.
    - Limitation types are matched by name. Limitation definitions are copied as they are, so a limitation that
      refers to objects by ID (e.g. a group) must be reviewed on the target.
    - User settings are copied, except the *ViewID settings (view IDs differ between the systems).
    - Orion accounts get the password from -InitialPassword (prompted once if not given). Passwords of the
      source are not readable and not migrated.
    - Not migrated: AccountSID / group info (created by the target), last-login history beyond LastLogin.

.PARAMETER SourceHost
    Hostname or IP of the source SolarWinds server.

.PARAMETER DestinationHost
    Hostname or IP of the target SolarWinds server.

.PARAMETER SourceCredential
    Credential for the source SWIS. Prompted if neither this nor -Trusted is given.

.PARAMETER DestinationCredential
    Credential for the target SWIS. Prompted if neither this nor -Trusted is given.

.PARAMETER Trusted
    Use a trusted (certificate) connection instead of credentials. Run the script as admin.

.PARAMETER NameFilter
    SWQL LIKE pattern for the AccountID. Default '%' (all). Handy to test with one account, e.g. 'CONTOSO\jdoe'.

.PARAMETER InitialPassword
    Password for migrated Orion accounts (type 1) as SecureString. Prompted once, when the first Orion
    account is created, if not given. Not used for Windows accounts.

.PARAMETER LogPath
    Log file. Default: next to the script, with timestamp.

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-Accounts-v2.0.ps1 -SourceHost old.contoso.local -DestinationHost new.contoso.local -WhatIf

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-Accounts-v2.0.ps1 -SourceHost 192.168.21.58 -DestinationHost 192.168.25.30 -NameFilter 'CONTOSO\%' -Verbose
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourceHost,
    [Parameter(Mandatory)][string]$DestinationHost,
    [pscredential]$SourceCredential,
    [pscredential]$DestinationCredential,
    [switch]$Trusted,
    [string]$NameFilter = '%',
    [securestring]$InitialPassword,
    [string]$LogPath = (Join-Path -Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) `
                                 -ChildPath ('SwisMigration_Accounts_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date)))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region State -----------------------------------------------------------------------------------

$script:Stats = [ordered]@{
    AccountsOnSource = 0; Created = 0; Skipped = 0; Failed = 0; LimitationsCreated = 0; SettingsCopied = 0
}
$script:Problems = [System.Collections.Generic.List[object]]::new()

# Y/N flags of Orion.Accounts that are sent to UpdateAccount as booleans
$script:FlagProperties = @(
    'Enabled', 'AllowNodeManagement', 'AllowMapManagement', 'AllowAdmin', 'CanClearEvents',
    'AllowReportManagement', 'AllowAlertManagement', 'AllowCustomize', 'AllowUnmanage',
    'AllowDisableAction', 'AllowDisableAlert', 'AllowDisableAllActions', 'DisableSessionTimeout'
)
# Properties copied 1:1 (only if they have a value on the source)
$script:ValueProperties = @(
    'AlertSound', 'MenuName', 'ReportFolder', 'AlertCategory', 'Expires', 'LastLogin', 'GroupPriority', 'AllowViewCopCheck'
)

#endregion

#region Helper functions ------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory, Position = 0)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1,-5}] {2}' -f (Get-Date), $Level, $Message
    try {
        # -WhatIf:$false - the log must also be written during a -WhatIf run
        Add-Content -Path $LogPath -Value $line -Encoding UTF8 -WhatIf:$false
    }
    catch {
        Write-Warning "Cannot write to log file ${LogPath}: $($_.Exception.Message)"
    }
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'DEBUG' { Write-Verbose $line }
        default { Write-Host $line }
    }
}

function Add-Problem {
    param([string]$Type, [string]$Name, [string]$Detail)
    $script:Problems.Add([pscustomobject]@{ Type = $Type; Name = $Name; Detail = $Detail })
}

function Test-IsEmpty {
    param($Value)
    return ($null -eq $Value) -or ($Value -is [DBNull]) -or ([string]$Value -eq '')
}

function ConvertTo-Id {
    # NULL / empty -> 0, otherwise the integer value.
    param($Value)
    if (Test-IsEmpty $Value) { return 0 }
    return [int]$Value
}

function ConvertTo-SqlLiteral {
    # Escapes a value for the T-SQL statements that go through Orion.Reporting.ExecuteSQL (no parameters there).
    param($Value)
    if (Test-IsEmpty $Value) { return 'NULL' }
    return "'" + ([string]$Value).Replace("'", "''") + "'"
}

function Invoke-Swql {
    # Thin wrapper around Get-SwisData: always parameterised queries, one place to add retry/logging.
    param(
        [Parameter(Mandatory)]$Swis,
        [Parameter(Mandatory)][string]$Query,
        [hashtable]$Parameters
    )
    if ($Parameters) { Get-SwisData $Swis $Query $Parameters }
    else             { Get-SwisData $Swis $Query }
}

function Invoke-OrionSql {
    # Raw SQL against the Orion database. Needed for properties that SWQL does not expose
    # (Accounts.SummaryViewID, Accounts.LimitationID1-3, inserting Limitations). Callers must use ConvertTo-SqlLiteral.
    param([Parameter(Mandatory)]$Swis, [Parameter(Mandatory)][string]$Sql)
    Invoke-SwisVerb $Swis 'Orion.Reporting' 'ExecuteSQL' $Sql
}

function Connect-Orion {
    param([string]$HostName, [pscredential]$Credential, [switch]$Trusted, [string]$Role)
    Write-Log "Connecting to $Role $HostName ..."
    try {
        if ($Trusted) { return Connect-Swis -Hostname $HostName -Trusted }
        if (-not $Credential) { $Credential = Get-Credential -Message "SolarWinds $Role ($HostName)" }
        return Connect-Swis -Hostname $HostName -Credential $Credential
    }
    catch {
        throw "Connection to $Role $HostName failed: $($_.Exception.Message)"
    }
}

function Get-InitialPasswordText {
    # Prompts once (lazily), only if an Orion account is really going to be created.
    if (-not $script:InitialPassword) {
        $script:InitialPassword = Read-Host -Prompt 'Initial password for migrated Orion accounts' -AsSecureString
    }
    return [System.Net.NetworkCredential]::new('', $script:InitialPassword).Password
}

#endregion

#region Account copy ----------------------------------------------------------------------------

function Resolve-TargetViewId {
    # Maps a source view to the target view with the same title + view group. Falls back to the default view (1).
    param($Source, $Target, $SourceViewId, [string]$Account, [string]$Purpose)

    $id = ConvertTo-Id $SourceViewId
    if ($id -le 1) { return 1 }

    $view = @(Invoke-Swql $Source 'SELECT ViewTitle, ViewGroupName FROM Orion.Views WHERE ViewID = @id' @{ id = $id }) | Select-Object -First 1
    if (-not $view) {
        Write-Log "Account '$Account': $Purpose (ViewID $id) not found on source, default view used" -Level WARN
        return 1
    }

    if (Test-IsEmpty $view.ViewGroupName) {
        $hit = @(Invoke-Swql $Target 'SELECT ViewID FROM Orion.Views WHERE ViewTitle = @t AND ViewGroupName IS NULL' @{ t = [string]$view.ViewTitle }) | Select-Object -First 1
    }
    else {
        $hit = @(Invoke-Swql $Target 'SELECT ViewID FROM Orion.Views WHERE ViewTitle = @t AND ViewGroupName = @g' @{ t = [string]$view.ViewTitle; g = [string]$view.ViewGroupName }) | Select-Object -First 1
    }
    if ($null -eq $hit) {
        Write-Log "Account '$Account': $Purpose '$($view.ViewTitle)' not found on target, default view used" -Level WARN
        return 1
    }
    return [int]$hit
}

function Get-SourceSummaryViewId {
    # SummaryViewID is not exposed by SWQL, so it is read from the Accounts table directly.
    param($Source, [string]$AccountId)
    $result = Invoke-OrionSql $Source ("SELECT SummaryViewID FROM [Accounts] WHERE AccountID = {0}" -f (ConvertTo-SqlLiteral $AccountId))
    if ($result -is [System.Xml.XmlNode]) {
        # SelectSingleNode instead of dotted access: a missing node must not throw under StrictMode
        $node = $result.SelectSingleNode("//*[local-name()='SummaryViewID']")
        if ($node -and $node.InnerText -match '^\d+$') { return [int]$node.InnerText }
    }
    return 1
}

function Get-AccountLimitations {
    # Reads the limitations (slot 1-3) of a source account and maps the limitation type to the target type ID.
    # Throws if a type does not exist on the target: creating the account without it would grant more access.
    param($Source, $Target, $Account)

    $limitations = [System.Collections.Generic.List[object]]::new()
    foreach ($slot in 1..3) {
        $limitId = ConvertTo-Id $Account."LimitationID$slot"
        if ($limitId -eq 0) { continue }

        $limit = @(Invoke-Swql $Source 'SELECT LimitationTypeID, Definition, WhereClause FROM Orion.Limitations WHERE LimitationID = @id' @{ id = $limitId }) | Select-Object -First 1
        if (-not $limit) { throw "Limitation $slot (LimitationID $limitId) not found on source." }

        $typeName = Invoke-Swql $Source 'SELECT Name FROM Orion.LimitationTypes WHERE LimitationTypeID = @id' @{ id = [int]$limit.LimitationTypeID }
        $typeId   = @(Invoke-Swql $Target 'SELECT LimitationTypeID FROM Orion.LimitationTypes WHERE Name = @n' @{ n = [string]$typeName }) | Select-Object -First 1
        if ($null -eq $typeId) { throw "Limitation type '$typeName' (slot $slot) does not exist on the target." }

        $limitations.Add([pscustomobject]@{
            Slot       = $slot
            TypeName   = [string]$typeName
            TypeId     = [int]$typeId
            Definition = $limit.Definition
            WhereClause = $limit.WhereClause
        })
    }
    return , $limitations.ToArray()
}

function New-TargetLimitation {
    # Inserts one limitation on the target and returns its ID.
    param($Target, $Limitation)
    $sql = 'INSERT INTO [Limitations] VALUES ({0}, {1}, {2})' -f $Limitation.TypeId, (ConvertTo-SqlLiteral $Limitation.Definition), (ConvertTo-SqlLiteral $Limitation.WhereClause)
    [void](Invoke-OrionSql $Target $sql)

    $newId = Invoke-Swql $Target 'SELECT TOP 1 LimitationID FROM Orion.Limitations ORDER BY LimitationID DESC'
    if (Test-IsEmpty $newId) { throw "Limitation '$($Limitation.TypeName)' was inserted, but its ID could not be read back." }
    $script:Stats.LimitationsCreated++
    return [int]$newId
}

function Get-AccountProperties {
    # Property bag for Orion.Accounts.UpdateAccount.
    param($Account, [int]$HomeViewId)

    $props = @{ HomePageViewID = $HomeViewId }
    foreach ($name in $script:FlagProperties) { $props[$name] = ([string]$Account.$name -eq 'Y') }
    foreach ($name in $script:ValueProperties) {
        if (-not (Test-IsEmpty $Account.$name)) { $props[$name] = $Account.$name }
    }
    return $props
}

function Copy-UserSettings {
    param($Source, $Target, [string]$AccountId)

    $settings = @(Invoke-Swql $Source "SELECT AccountID, SettingName, SettingValue, InstanceSiteId FROM Orion.UserSettings WHERE AccountID = @a AND SettingName NOT LIKE '%ViewID'" @{ a = $AccountId })
    foreach ($s in $settings) {
        try {
            [void](New-SwisObject $Target -EntityType 'Orion.UserSettings' -Properties @{
                AccountID      = $s.AccountID
                SettingName    = $s.SettingName
                SettingValue   = $s.SettingValue
                InstanceSiteId = $s.InstanceSiteId
            })
            $script:Stats.SettingsCopied++
        }
        catch {
            # A single setting is not worth failing the account
            Write-Log "Account '$AccountId': setting '$($s.SettingName)' could not be copied: $($_.Exception.Message)" -Level WARN
            Add-Problem 'SettingFailed' $AccountId "$($s.SettingName): $($_.Exception.Message)"
        }
    }
}

function Remove-TargetAccount {
    # Removes a just-created, half-configured account so that a re-run can retry it. Returns $true on success.
    param($Target, [string]$AccountId)
    try {
        $uri = @(Invoke-Swql $Target 'SELECT Uri FROM Orion.Accounts WHERE AccountID = @a' @{ a = $AccountId }) | Select-Object -First 1
        if ($uri) { Remove-SwisObject $Target ([string]$uri) }
        return $true
    }
    catch {
        Write-Log "Removing half-configured account '$AccountId' failed: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Copy-Account {
    [CmdletBinding(SupportsShouldProcess)]
    param($Source, $Target, $Account, [System.Collections.Generic.HashSet[string]]$TargetAccountIds)

    $id   = [string]$Account.AccountID
    $type = ConvertTo-Id $Account.AccountType

    if ($type -notin 1, 2, 3) {
        Write-Log "Account '$id': account type $type is not supported, skipped" -Level WARN
        $script:Stats.Skipped++
        return
    }
    if ($TargetAccountIds.Contains($id)) {
        Write-Log "Account '$id' already exists on target, skipped"
        $script:Stats.Skipped++
        return
    }

    # 1. read-only preparation - everything that can fail without touching the target
    $homeViewId    = Resolve-TargetViewId -Source $Source -Target $Target -SourceViewId $Account.HomePageViewID -Account $id -Purpose 'Home page view'
    $summaryViewId = Resolve-TargetViewId -Source $Source -Target $Target -SourceViewId (Get-SourceSummaryViewId -Source $Source -AccountId $id) -Account $id -Purpose 'Summary view'
    $limitations   = @(Get-AccountLimitations -Source $Source -Target $Target -Account $Account)

    if (-not $PSCmdlet.ShouldProcess($id, "Create account (type $type, $($limitations.Count) limitation(s)) on $DestinationHost")) { return }

    # 2. create the account
    Write-Log "Migrating account '$id' (type $type)"
    if ($type -eq 1) {
        [void](Invoke-SwisVerb $Target 'Orion.Accounts' 'CreateOrionAccount' @($id, (Get-InitialPasswordText)))
    }
    else {
        [void](Invoke-SwisVerb $Target 'Orion.Accounts' 'CreateWindowsAccount' @($type, $id))
    }
    [void]$TargetAccountIds.Add($id)

    # 3. configure it - any failure removes the account again
    try {
        $limitIds = @{ 1 = 0; 2 = 0; 3 = 0 }
        foreach ($limitation in $limitations) { $limitIds[$limitation.Slot] = New-TargetLimitation -Target $Target -Limitation $limitation }

        [void](Invoke-SwisVerb $Target 'Orion.Accounts' 'UpdateAccount' @($id, (Get-AccountProperties -Account $Account -HomeViewId $homeViewId)))

        # Columns that UpdateAccount does not cover
        $sql = 'UPDATE [Accounts] SET LimitationID1 = {0}, LimitationID2 = {1}, LimitationID3 = {2}, SummaryViewID = {3} WHERE AccountID = {4}' -f `
               $limitIds[1], $limitIds[2], $limitIds[3], $summaryViewId, (ConvertTo-SqlLiteral $id)
        [void](Invoke-OrionSql $Target $sql)

        Copy-UserSettings -Source $Source -Target $Target -AccountId $id
    }
    catch {
        $message = $_.Exception.Message
        if (Remove-TargetAccount -Target $Target -AccountId $id) {
            [void]$TargetAccountIds.Remove($id)
            throw "Configuration failed, account removed from target again: $message"
        }
        throw "Configuration failed, account is half-configured on the target - remove it manually before re-running: $message"
    }

    $script:Stats.Created++
    Write-Log "Created account '$id'"
}

function Copy-Accounts {
    [CmdletBinding(SupportsShouldProcess)]
    param($Source, $Target)

    $targetIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in @(Invoke-Swql $Target 'SELECT AccountID FROM Orion.Accounts')) { [void]$targetIds.Add([string]$a) }
    Write-Log "Target has $($targetIds.Count) account(s)"

    $accounts = @(Invoke-Swql $Source @'
SELECT AccountID, AccountType, Enabled, AllowNodeManagement, AllowMapManagement, AllowAdmin, CanClearEvents,
       AllowReportManagement, AllowAlertManagement, AllowCustomize, AllowUnmanage, AllowDisableAction,
       AllowDisableAlert, AllowDisableAllActions, AlertSound, MenuName, DisableSessionTimeout, ReportFolder,
       AlertCategory, Expires, LastLogin, GroupPriority, AllowViewCopCheck,
       LimitationID1, LimitationID2, LimitationID3, HomePageViewID
FROM Orion.Accounts
WHERE AccountID LIKE @f
ORDER BY AccountID
'@ @{ f = $NameFilter })
    $script:Stats.AccountsOnSource = $accounts.Count
    Write-Log "$($accounts.Count) account(s) on source (filter '$NameFilter')"

    foreach ($account in $accounts) {
        try {
            Copy-Account -Source $Source -Target $Target -Account $account -TargetAccountIds $targetIds
        }
        catch {
            Write-Log "Account '$($account.AccountID)' FAILED: $($_.Exception.Message)" -Level ERROR
            Write-Log $_.ScriptStackTrace -Level DEBUG
            $script:Stats.Failed++
            Add-Problem 'AccountFailed' ([string]$account.AccountID) $_.Exception.Message
        }
    }
}

#endregion

#region Main ------------------------------------------------------------------------------------

Write-Log "Account migration started: $SourceHost -> $DestinationHost (WhatIf: $([bool]$WhatIfPreference))"
Write-Log "Log file: $LogPath"

$script:InitialPassword = $InitialPassword

try {
    if ($SourceHost -eq $DestinationHost) { throw 'Source and destination are the same host.' }

    $sourceSwis = Connect-Orion -HostName $SourceHost      -Credential $SourceCredential      -Trusted:$Trusted -Role 'source'
    $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'

    Copy-Accounts -Source $sourceSwis -Target $targetSwis

    if ($script:Stats.Created -gt 0) {
        Write-Log 'Review limitations on the target - definitions that refer to objects by ID (groups, custom property values) are copied as they are.' -Level WARN
    }
}
catch {
    Write-Log "Migration aborted: $($_.Exception.Message)" -Level ERROR
    Add-Problem 'Fatal' 'Script' $_.Exception.Message
}
finally {
    Write-Log '----- Summary -----'
    foreach ($key in $script:Stats.Keys) { Write-Log ('{0,-20} {1}' -f $key, $script:Stats[$key]) }
    if ($script:Problems.Count -gt 0) {
        Write-Log "$($script:Problems.Count) problem(s):" -Level WARN
        foreach ($problem in $script:Problems) { Write-Log "  [$($problem.Type)] $($problem.Name): $($problem.Detail)" -Level WARN }
    }
    Write-Log 'Done.'
}

if ($script:Problems.Count -gt 0) { exit 1 }

#endregion
