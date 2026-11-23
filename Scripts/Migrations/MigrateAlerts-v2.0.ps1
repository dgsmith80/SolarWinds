#Requires -Version 5.1
#Requires -Modules SwisPowerShell
<#
.SYNOPSIS
    Migrates custom (non-canned) alert definitions from one SolarWinds Platform (Orion) instance
    to another via the Orion.AlertConfigurations Export/Import verbs.

.DESCRIPTION
    Author:     Holger Mundt
    Email:      holger.mundt@loop1.com
    Version:    2.0
    History:    2025-03-11 HM - Creation of Script (v1.0)
                v2.0    Refactoring of v1.0:
                          - no hard-coded credentials / IPs (parameters, Get-Credential or -Trusted)
                          - failure to read the source alerts now aborts (v1.0 continued with an empty list)
                          - the alert is exported only if it is going to be imported (v1.0 exported before the exists-check)
                          - import result is evaluated robustly and verified against the target afterwards
                          - a half-imported alert is removed again, so a re-run can retry it
                          - every created alert is written to a manifest; -Rollback removes exactly those alerts
                          - every exported alert is stored as XML file (backup / manual re-import)
                          - -WhatIf support, log file, summary, exit code
                          - pre-flight check for custom properties missing on the target
                          - optional -DisableAfterImport to avoid double notifications

    What is NOT migrated: canned (built-in) alerts, including modified ones. Alert actions that reference
    the old environment (SMTP server, scripts/paths, credentials, URLs) are imported as they are and must
    be reviewed on the target.

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
    SWQL LIKE pattern for the alert name. Default '%' (all). Handy to test with one alert, e.g. 'Test%'.

.PARAMETER DisableAfterImport
    Disable each alert on the target right after the import, so that source and target do not both send
    notifications. Enable the alerts on the target after review.

.PARAMETER Rollback
    Removes the alerts listed in -ManifestPath from the target. Only alerts whose ID and name still match
    the manifest are removed. Nothing is read from or changed on the source.

.PARAMETER ManifestPath
    CSV with the alerts created on the target (written by a normal run, read by -Rollback).
    Default: next to the script, with timestamp. Mandatory together with -Rollback.

.PARAMETER ExportPath
    Folder for the exported alert XML files. Default: next to the script, with timestamp.

.PARAMETER LogPath
    Log file. Default: next to the script, with timestamp.

.EXAMPLE
    .\MigrateAlerts-v2.0.ps1 -SourceHost 192.168.50.106 -DestinationHost 192.168.11.20 -WhatIf

.EXAMPLE
    .\MigrateAlerts-v2.0.ps1 -SourceHost old.contoso.local -DestinationHost new.contoso.local -DisableAfterImport

.EXAMPLE
    .\MigrateAlerts-v2.0.ps1 -DestinationHost new.contoso.local -SourceHost old.contoso.local -Rollback -ManifestPath .\AlertManifest_20261029_210000.csv
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourceHost,
    [Parameter(Mandatory)][string]$DestinationHost,
    [pscredential]$SourceCredential,
    [pscredential]$DestinationCredential,
    [switch]$Trusted,
    [string]$NameFilter = '%',
    [switch]$DisableAfterImport,
    [switch]$Rollback,
    [string]$ManifestPath,
    [string]$ExportPath,
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Defaults ---------------------------------------------------------------------------------

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$Stamp     = '{0:yyyyMMdd_HHmmss}' -f (Get-Date)
if (-not $LogPath)      { $LogPath      = Join-Path $ScriptDir "AlertMigration_$Stamp.log" }
if (-not $ExportPath)   { $ExportPath   = Join-Path $ScriptDir "AlertExport_$Stamp" }
if (-not $ManifestPath) {
    if ($Rollback) { throw '-Rollback needs -ManifestPath (the CSV written by the migration run).' }
    $ManifestPath = Join-Path $ScriptDir "AlertManifest_$Stamp.csv"
}

$script:Stats    = [ordered]@{ AlertsOnSource = 0; Created = 0; Skipped = 0; Failed = 0; Disabled = 0; RolledBack = 0 }
$script:Problems = [System.Collections.Generic.List[object]]::new()

#endregion

#region Helper functions ------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory, Position = 0)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1,-5}] {2}' -f (Get-Date), $Level, $Message
    try   { Add-Content -Path $LogPath -Value $line -Encoding UTF8 -WhatIf:$false }   # also log during -WhatIf
    catch { Write-Warning "Cannot write to log file ${LogPath}: $($_.Exception.Message)" }
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

function Invoke-Swql {
    # Thin wrapper around Get-SwisData: always parameterised queries.
    param(
        [Parameter(Mandatory)]$Swis,
        [Parameter(Mandatory)][string]$Query,
        [hashtable]$Parameters
    )
    if ($Parameters) { Get-SwisData $Swis $Query $Parameters }
    else             { Get-SwisData $Swis $Query }
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

function Get-VerbResultText {
    # Invoke-SwisVerb returns an XmlElement (or nothing). InnerText also works if the element has child nodes,
    # and it does not throw under StrictMode when a child such as MigrationMessage is missing.
    param($Result)
    if ($null -eq $Result) { return '' }
    if ($Result -is [System.Xml.XmlNode]) { return [string]$Result.InnerText }
    return [string]$Result
}

function Get-SafeFileName {
    param([string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    return ($Name -replace "[$([regex]::Escape($invalid))]", '_')
}

function Add-ManifestEntry {
    param([int]$AlertId, [string]$Name)
    [pscustomobject]@{
        TargetAlertID = $AlertId
        Name          = $Name
        CreatedAt     = (Get-Date).ToString('s')
        Host          = $DestinationHost
    } | Export-Csv -Path $ManifestPath -Append -NoTypeInformation -Encoding UTF8 -WhatIf:$false
}

function Remove-TargetAlert {
    # Removes one alert on the target by Uri. Returns $true on success.
    param($Target, [string]$Uri)
    try   { Remove-SwisObject $Target $Uri; return $true }
    catch { Write-Log "Removing $Uri failed: $($_.Exception.Message)" -Level ERROR; return $false }
}

#endregion

#region Pre-flight ------------------------------------------------------------------------------

function Test-TargetCustomProperties {
    # Alert conditions/scopes often use custom properties. If they are missing on the target the import fails or
    # the alert is incomplete. Only warns - the alerts that do not need them will import fine.
    param($Source, $Target)
    $have = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($cp in @(Invoke-Swql $Target 'SELECT Table, Field FROM Orion.CustomProperty')) { [void]$have.Add("$($cp.Table).$($cp.Field)") }

    $missing = @(foreach ($cp in @(Invoke-Swql $Source 'SELECT Table, Field FROM Orion.CustomProperty')) {
        if (-not $have.Contains("$($cp.Table).$($cp.Field)")) { "$($cp.Table).$($cp.Field)" }
    })
    if ($missing.Count -gt 0) {
        Write-Log "$($missing.Count) custom property(ies) exist on the source but not on the target: $($missing -join ', ')" -Level WARN
        Write-Log 'Alerts that use them may fail to import - run the platform migration script (custom properties) first.' -Level WARN
    }
}

#endregion

#region Alert copy ------------------------------------------------------------------------------

function Copy-SwisAlert {
    [CmdletBinding(SupportsShouldProcess)]
    param($Source, $Target, $Alert)

    $name = [string]$Alert.Name
    $uriQuery = 'SELECT AlertID, Uri FROM Orion.AlertConfigurations WHERE Name = @n'

    # 1. already there? (cheap check first, export only what is actually needed)
    if (@(Invoke-Swql $Target $uriQuery @{ n = $name }).Count -gt 0) {
        Write-Log "Skipping '$name' - already exists on $DestinationHost"
        $script:Stats.Skipped++
        return
    }

    # 2. export from source and keep a copy on disk
    $exportResult = Invoke-SwisVerb $Source 'Orion.AlertConfigurations' 'Export' @([int]$Alert.AlertID, $false)
    $alertXml     = Get-VerbResultText $exportResult
    if (Test-IsEmpty $alertXml) { throw 'Export returned no data.' }

    $file = Join-Path $ExportPath ('{0}_{1}.xml' -f $Alert.AlertID, (Get-SafeFileName $name))
    Set-Content -Path $file -Value $alertXml -Encoding UTF8 -WhatIf:$false
    Write-Log "Exported '$name' -> $file" -Level DEBUG

    # 3. import into target
    if (-not $PSCmdlet.ShouldProcess($name, "Import alert on $DestinationHost")) { return }
    Write-Log "Creating alert '$name' on $DestinationHost"

    $importText = Get-VerbResultText (Invoke-SwisVerb $Target 'Orion.AlertConfigurations' 'Import' @($alertXml, $false))
    $created    = @(Invoke-Swql $Target $uriQuery @{ n = $name })

    if ($importText -match 'failed with error') {
        # The import can leave a half-created alert behind; a re-run would then skip it as "already exists".
        foreach ($c in $created) {
            if (Remove-TargetAlert -Target $Target -Uri ([string]$c.Uri)) { $importText += ' | partial alert removed from target again' }
            else { $importText += " | ROLLBACK FAILED, remove manually: $($c.Uri)" }
        }
        throw "Import failed: $importText"
    }
    if ($created.Count -eq 0) {
        throw "Import reported no error, but the alert was not found on the target afterwards. Result: '$importText'"
    }

    $newId = [int]$created[0].AlertID
    Add-ManifestEntry -AlertId $newId -Name $name
    $script:Stats.Created++
    Write-Log "Created alert '$name' (AlertID $newId)"

    # 4. optional: disable. A failure here is a warning, not a failed alert.
    if ($DisableAfterImport) {
        try {
            Set-SwisObject $Target ([string]$created[0].Uri) @{ Enabled = $false }
            $script:Stats.Disabled++
        }
        catch {
            Write-Log "Alert '$name' created but could NOT be disabled: $($_.Exception.Message)" -Level WARN
            Add-Problem 'NotDisabled' $name $_.Exception.Message
        }
    }
}

function Undo-AlertMigration {
    [CmdletBinding(SupportsShouldProcess)]
    param($Target)

    if (-not (Test-Path -Path $ManifestPath)) { throw "Manifest not found: $ManifestPath" }
    $entries = @(Import-Csv -Path $ManifestPath)
    Write-Log "Rollback: $($entries.Count) alert(s) in $ManifestPath"

    foreach ($e in $entries) {
        try {
            # ID and name must both match - protects against removing an alert that was recreated by someone else
            $uri = @(Invoke-Swql $Target 'SELECT Uri FROM Orion.AlertConfigurations WHERE AlertID = @id AND Name = @n' `
                                 @{ id = [int]$e.TargetAlertID; n = [string]$e.Name }) | Select-Object -First 1
            if (-not $uri) {
                Write-Log "Rollback: '$($e.Name)' (AlertID $($e.TargetAlertID)) not found on target, nothing to do"
                continue
            }
            if ($PSCmdlet.ShouldProcess($e.Name, "Remove alert (AlertID $($e.TargetAlertID)) from $DestinationHost")) {
                if (Remove-TargetAlert -Target $Target -Uri ([string]$uri)) {
                    $script:Stats.RolledBack++
                    Write-Log "Rollback: removed '$($e.Name)'"
                }
                else {
                    $script:Stats.Failed++
                    Add-Problem 'RollbackFailed' ([string]$e.Name) 'See log'
                }
            }
        }
        catch {
            Write-Log "Rollback of '$($e.Name)' FAILED: $($_.Exception.Message)" -Level ERROR
            $script:Stats.Failed++
            Add-Problem 'RollbackFailed' ([string]$e.Name) $_.Exception.Message
        }
    }
}

#endregion

#region Main ------------------------------------------------------------------------------------

Write-Log "Started: $SourceHost -> $DestinationHost (Rollback: $([bool]$Rollback), WhatIf: $([bool]$WhatIfPreference), DisableAfterImport: $([bool]$DisableAfterImport))"
Write-Log "Log file: $LogPath"

try {
    if ($SourceHost -eq $DestinationHost) { throw 'Source and destination are the same host.' }

    if ($Rollback) {
        $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'
        Undo-AlertMigration -Target $targetSwis
    }
    else {
        $sourceSwis = Connect-Orion -HostName $SourceHost      -Credential $SourceCredential      -Trusted:$Trusted -Role 'source'
        $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'

        if (-not (Test-Path -Path $ExportPath)) { [void](New-Item -ItemType Directory -Path $ExportPath -Force -WhatIf:$false) }

        Test-TargetCustomProperties -Source $sourceSwis -Target $targetSwis

        # All custom alerts - a failure here is fatal (v1.0 carried on with an undefined list)
        $alerts = @(Invoke-Swql $sourceSwis 'SELECT AlertID, Name FROM Orion.AlertConfigurations WHERE Canned = FALSE AND Name LIKE @f ORDER BY Name' @{ f = $NameFilter })
        $script:Stats.AlertsOnSource = $alerts.Count
        Write-Log "$($alerts.Count) custom alert(s) on source (filter '$NameFilter')"

        foreach ($alert in $alerts) {
            try {
                Copy-SwisAlert -Source $sourceSwis -Target $targetSwis -Alert $alert
            }
            catch {
                Write-Log "Alert '$($alert.Name)' FAILED: $($_.Exception.Message)" -Level ERROR
                $script:Stats.Failed++
                Add-Problem 'AlertFailed' ([string]$alert.Name) $_.Exception.Message
            }
        }

        if ($script:Stats.Created -gt 0) {
            Write-Log "Manifest for rollback: $ManifestPath"
            Write-Log 'Review alert actions on the target (SMTP server, recipients, scripts, URLs, credentials) before enabling.' -Level WARN
        }
    }
}
catch {
    Write-Log "Aborted: $($_.Exception.Message)" -Level ERROR
    Add-Problem 'Fatal' 'Script' $_.Exception.Message
}
finally {
    Write-Log '----- Summary -----'
    foreach ($key in $script:Stats.Keys) { Write-Log ('{0,-16} {1}' -f $key, $script:Stats[$key]) }
    if ($script:Problems.Count -gt 0) {
        Write-Log "$($script:Problems.Count) problem(s):" -Level WARN
        foreach ($p in $script:Problems) { Write-Log "  [$($p.Type)] $($p.Name): $($p.Detail)" -Level WARN }
    }
    Write-Log 'Done.'
}

if ($script:Problems.Count -gt 0) { exit 1 }

#endregion
