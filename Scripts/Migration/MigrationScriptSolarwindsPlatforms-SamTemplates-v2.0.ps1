#Requires -Version 5.1
#Requires -Modules SwisPowerShell
<#
.SYNOPSIS
    Migrates SAM application templates from one SolarWinds Platform (Orion) instance to another via
    the Orion.APM.ApplicationTemplate ExportTemplate / ImportTemplate verbs.

.DESCRIPTION
    Author:     Holger Mundt
    Email:      holger.mundt@loop1.com
    Version:    2.0
    History:    2024-02-11 HM - Creation of Script (v1.0)
                v2.0    Refactoring of v1.0:
                          - no hard-coded credentials / hosts (parameters, Get-Credential or -Trusted)
                          - parameterised SWQL via Invoke-Swql, same helper functions as the other v2.0 scripts
                          - failure to read the source templates now aborts (v1.0 continued with an empty list)
                          - exists-check by name is a cheap query first, export only what is actually imported
                          - import result is verified against the target afterwards (v1.0 only checked the verb result)
                          - every created template is written to a manifest; -Rollback removes exactly those templates
                          - every exported template is stored as XML file (backup / manual re-import)
                          - -WhatIf support, log file, summary, exit code

    Runs standalone, independent of the CP + Nodes and Groups scripts. Assigning the templates to nodes
    (applications) is NOT part of this script.

    What is NOT migrated: credentials that templates reference (credential-based component monitors need
    a credential with the same name on the target) and assigned applications. Templates that already
    exist on the target (same name, e.g. the built-in ones) are skipped, not updated.

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
    SWQL LIKE pattern for the template name. Default '%' (all). Handy to test with one template, e.g. 'Test%'.

.PARAMETER Rollback
    Removes the templates listed in -ManifestPath from the target. Only templates whose ID and name still
    match the manifest are removed. Nothing is read from or changed on the source.

.PARAMETER ManifestPath
    CSV with the templates created on the target (written by a normal run, read by -Rollback).
    Default: next to the script, with timestamp. Mandatory together with -Rollback.

.PARAMETER ExportPath
    Folder for the exported template XML files. Default: next to the script, with timestamp.

.PARAMETER LogPath
    Log file. Default: next to the script, with timestamp.

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-SamTemplates-v2.0.ps1 -SourceHost old.contoso.local -DestinationHost new.contoso.local -WhatIf

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-SamTemplates-v2.0.ps1 -SourceHost 192.168.50.106 -DestinationHost 192.168.2.200 -NameFilter 'Loop1%' -Verbose

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-SamTemplates-v2.0.ps1 -SourceHost old.contoso.local -DestinationHost new.contoso.local -Rollback -ManifestPath .\SamTemplateManifest_20261029_210000.csv
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourceHost,
    [Parameter(Mandatory)][string]$DestinationHost,
    [pscredential]$SourceCredential,
    [pscredential]$DestinationCredential,
    [switch]$Trusted,
    [string]$NameFilter = '%',
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
if (-not $LogPath)      { $LogPath    = Join-Path $ScriptDir "SwisMigration_SamTemplates_$Stamp.log" }
if (-not $ExportPath)   { $ExportPath = Join-Path $ScriptDir "SamTemplateExport_$Stamp" }
if (-not $ManifestPath) {
    if ($Rollback) { throw '-Rollback needs -ManifestPath (the CSV written by the migration run).' }
    $ManifestPath = Join-Path $ScriptDir "SamTemplateManifest_$Stamp.csv"
}

#endregion

#region State -----------------------------------------------------------------------------------

$script:Stats = [ordered]@{
    TemplatesOnSource = 0; TemplatesCreated = 0; TemplatesSkipped = 0; TemplatesFailed = 0; RolledBack = 0
}
$script:Problems = [System.Collections.Generic.List[object]]::new()

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
    # and it does not throw under StrictMode when a child is missing.
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
    param([int]$TemplateId, [string]$Name)
    [pscustomobject]@{
        TargetTemplateID = $TemplateId
        Name             = $Name
        CreatedAt        = (Get-Date).ToString('s')
        Host             = $DestinationHost
    } | Export-Csv -Path $ManifestPath -Append -NoTypeInformation -Encoding UTF8 -WhatIf:$false
}

function Remove-TargetTemplate {
    # Removes one template on the target by Uri. Returns $true on success.
    param($Target, [string]$Uri)
    try   { Remove-SwisObject $Target $Uri; return $true }
    catch { Write-Log "Removing $Uri failed: $($_.Exception.Message)" -Level ERROR; return $false }
}

#endregion

#region Pre-flight ------------------------------------------------------------------------------

function Test-SamAvailable {
    # SAM must be installed on both sides, otherwise the entity does not exist.
    param($Swis, [string]$Role)
    try   { [void](Invoke-Swql $Swis 'SELECT TOP 1 ApplicationTemplateID FROM Orion.APM.ApplicationTemplate') }
    catch { throw "SAM (Orion.APM.ApplicationTemplate) is not available on the $Role: $($_.Exception.Message)" }
}

#endregion

#region Template copy ---------------------------------------------------------------------------

function Copy-SamTemplate {
    [CmdletBinding(SupportsShouldProcess)]
    param($Source, $Target, $Template)

    $name     = [string]$Template.Name
    $findQuery = 'SELECT ApplicationTemplateID, Uri FROM Orion.APM.ApplicationTemplate WHERE Name = @n'

    # 1. already there? (cheap check first, export only what is actually needed)
    if (@(Invoke-Swql $Target $findQuery @{ n = $name }).Count -gt 0) {
        Write-Log "Skipping '$name' - already exists on $DestinationHost"
        $script:Stats.TemplatesSkipped++
        return
    }

    # 2. export from source and keep a copy on disk
    $templateXml = Get-VerbResultText (Invoke-SwisVerb $Source 'Orion.APM.ApplicationTemplate' 'ExportTemplate' @([int]$Template.ApplicationTemplateID))
    if (Test-IsEmpty $templateXml) { throw 'Export returned no data.' }

    $file = Join-Path $ExportPath ('{0}_{1}.xml' -f $Template.ApplicationTemplateID, (Get-SafeFileName $name))
    Set-Content -Path $file -Value $templateXml -Encoding UTF8 -WhatIf:$false
    Write-Log "Exported '$name' -> $file" -Level DEBUG

    # 3. import into target
    if (-not $PSCmdlet.ShouldProcess($name, "Import SAM template on $DestinationHost")) { return }
    Write-Log "Creating SAM template '$name' on $DestinationHost"

    $importText = Get-VerbResultText (Invoke-SwisVerb $Target 'Orion.APM.ApplicationTemplate' 'ImportTemplate' @($templateXml))

    # 4. verify: the verb result alone is not reliable, look for the template on the target
    $created = @(Invoke-Swql $Target $findQuery @{ n = $name })
    if ($created.Count -eq 0) {
        throw "Import reported no error, but the template was not found on the target afterwards. Result: '$importText'"
    }

    $newId = [int]$created[0].ApplicationTemplateID
    Add-ManifestEntry -TemplateId $newId -Name $name
    $script:Stats.TemplatesCreated++
    Write-Log "Created SAM template '$name' (ApplicationTemplateID $newId)"
}

function Undo-SamTemplateMigration {
    [CmdletBinding(SupportsShouldProcess)]
    param($Target)

    if (-not (Test-Path -Path $ManifestPath)) { throw "Manifest not found: $ManifestPath" }
    $entries = @(Import-Csv -Path $ManifestPath)
    Write-Log "Rollback: $($entries.Count) template(s) in $ManifestPath"

    foreach ($e in $entries) {
        try {
            # ID and name must both match - protects against removing a template that was recreated by someone else
            $uri = @(Invoke-Swql $Target 'SELECT Uri FROM Orion.APM.ApplicationTemplate WHERE ApplicationTemplateID = @id AND Name = @n' `
                                 @{ id = [int]$e.TargetTemplateID; n = [string]$e.Name }) | Select-Object -First 1
            if (-not $uri) {
                Write-Log "Rollback: '$($e.Name)' (ApplicationTemplateID $($e.TargetTemplateID)) not found on target, nothing to do"
                continue
            }
            if ($PSCmdlet.ShouldProcess($e.Name, "Remove SAM template (ApplicationTemplateID $($e.TargetTemplateID)) from $DestinationHost")) {
                if (Remove-TargetTemplate -Target $Target -Uri ([string]$uri)) {
                    $script:Stats.RolledBack++
                    Write-Log "Rollback: removed '$($e.Name)'"
                }
                else {
                    $script:Stats.TemplatesFailed++
                    Add-Problem 'RollbackFailed' ([string]$e.Name) 'See log'
                }
            }
        }
        catch {
            Write-Log "Rollback of '$($e.Name)' FAILED: $($_.Exception.Message)" -Level ERROR
            $script:Stats.TemplatesFailed++
            Add-Problem 'RollbackFailed' ([string]$e.Name) $_.Exception.Message
        }
    }
}

#endregion

#region Main ------------------------------------------------------------------------------------

Write-Log "SAM template migration started: $SourceHost -> $DestinationHost (Rollback: $([bool]$Rollback), WhatIf: $([bool]$WhatIfPreference))"
Write-Log "Log file: $LogPath"

try {
    if ($SourceHost -eq $DestinationHost) { throw 'Source and destination are the same host.' }

    if ($Rollback) {
        $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'
        Undo-SamTemplateMigration -Target $targetSwis
    }
    else {
        $sourceSwis = Connect-Orion -HostName $SourceHost      -Credential $SourceCredential      -Trusted:$Trusted -Role 'source'
        $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'

        Test-SamAvailable -Swis $sourceSwis -Role 'source'
        Test-SamAvailable -Swis $targetSwis -Role 'destination'

        if (-not (Test-Path -Path $ExportPath)) { [void](New-Item -ItemType Directory -Path $ExportPath -Force -WhatIf:$false) }

        # All templates - a failure here is fatal (v1.0 carried on with an undefined list)
        $templates = @(Invoke-Swql $sourceSwis 'SELECT ApplicationTemplateID, Name FROM Orion.APM.ApplicationTemplate WHERE Name LIKE @f ORDER BY Name' @{ f = $NameFilter })
        $script:Stats.TemplatesOnSource = $templates.Count
        Write-Log "$($templates.Count) SAM template(s) on source (filter '$NameFilter')"

        foreach ($template in $templates) {
            try {
                Copy-SamTemplate -Source $sourceSwis -Target $targetSwis -Template $template
            }
            catch {
                Write-Log "SAM template '$($template.Name)' FAILED: $($_.Exception.Message)" -Level ERROR
                Write-Log $_.ScriptStackTrace -Level DEBUG
                $script:Stats.TemplatesFailed++
                Add-Problem 'TemplateFailed' ([string]$template.Name) $_.Exception.Message
            }
        }

        if ($script:Stats.TemplatesCreated -gt 0) {
            Write-Log "Manifest for rollback: $ManifestPath"
            Write-Log 'Credentials used by the templates (component monitors) must exist on the target with the same name.' -Level WARN
        }
    }
}
catch {
    Write-Log "Migration aborted: $($_.Exception.Message)" -Level ERROR
    Add-Problem 'Fatal' 'Script' $_.Exception.Message
}
finally {
    Write-Log '----- Summary -----'
    foreach ($key in $script:Stats.Keys) { Write-Log ('{0,-24} {1}' -f $key, $script:Stats[$key]) }
    if ($script:Problems.Count -gt 0) {
        Write-Log "$($script:Problems.Count) problem(s):" -Level WARN
        foreach ($problem in $script:Problems) { Write-Log "  [$($problem.Type)] $($problem.Name): $($problem.Detail)" -Level WARN }
    }
    Write-Log 'Done.'
}

if ($script:Problems.Count -gt 0) { exit 1 }

#endregion
