#Requires -Version 5.1
#Requires -Modules SwisPowerShell
<#
.SYNOPSIS
    Step 2 of 2: migrates groups (including the parent/child hierarchy and their members) from one
    SolarWinds Platform (Orion) instance to another via SWIS.

.DESCRIPTION
    Author:     Holger Mundt
    Email:      holger.mundt@loop1.com
    Version:    2.0
    History:    v2.0    Split from MigrationScriptSolarwindsPlatforms-v2.0.ps1 (group part)

    Runs standalone, but must be run AFTER MigrationScriptSolarwindsPlatforms-CPNodes-v2.0.ps1:
    node and interface members are matched to the nodes/interfaces that already exist on the target
    (nodes by IP address, then caption; interfaces by full name). Members that cannot be found are
    skipped with a warning; re-running the script does NOT repair them, because existing groups are
    skipped. Run the CP + Nodes script first and completely.

    - Groups are processed parent-first, the parent is found by exact ContainerID (no LIKE '%...%')
    - Existing groups (same name) on the target are skipped
    - A group needs at least one member, so groups without resolvable members get a dummy filter
      ("Dummy Devices", Vendor='Dummy') as in v1.7
    - Not migrated: members of type application or volume (warning), group custom property values

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

.PARAMETER LogPath
    Log file. Default: next to the script, with timestamp.

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-Groups-v2.0.ps1 -SourceHost old.contoso.local -DestinationHost new.contoso.local -WhatIf

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-Groups-v2.0.ps1 -SourceHost 192.168.50.106 -DestinationHost 192.168.2.200 -Verbose
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourceHost,
    [Parameter(Mandatory)][string]$DestinationHost,
    [pscredential]$SourceCredential,
    [pscredential]$DestinationCredential,
    [switch]$Trusted,
    [string]$LogPath = (Join-Path -Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) `
                                 -ChildPath ('SwisMigration_Groups_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date)))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region State -----------------------------------------------------------------------------------

$script:Stats = [ordered]@{
    GroupsCreated = 0; GroupsSkipped = 0; GroupsFailed = 0
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

#endregion

#region Group copy ------------------------------------------------------------------------------

function Resolve-ContainerMembers {
    # Translates source group member definitions to target ones (nodes by IP, interfaces by FullName).
    param($Source, $Target, $Members, [string]$GroupName)
    $resolved = [System.Collections.Generic.List[object]]::new()

    foreach ($m in @($Members)) {
        if ($null -eq $m) { continue }
        try {
            if ([string]$m.Definition -like 'filter*') {
                $resolved.Add(@{ Name = [string]$m.Name; Definition = [string]$m.Definition })
                continue
            }
            switch ([string]$m.Entity) {
                'Orion.Nodes' {
                    $srcNode = Get-SwisObject $Source ([string]$m.Definition)
                    $hit = @(Invoke-Swql $Target 'SELECT Uri, NodeID FROM Orion.Nodes WHERE IPAddress = @ip' @{ ip = $srcNode.IPAddress }) | Select-Object -First 1
                    if (-not $hit) {
                        $hit = @(Invoke-Swql $Target 'SELECT Uri, NodeID FROM Orion.Nodes WHERE Caption = @c' @{ c = $srcNode.Caption }) | Select-Object -First 1
                    }
                    if ($hit) { $resolved.Add(@{ Name = "Orion.Nodes.NodeID=$($hit.NodeID)"; Definition = [string]$hit.Uri }) }
                    else      { Write-Log "Group '$GroupName': node $($srcNode.Caption) ($($srcNode.IPAddress)) not found on target, member skipped" -Level WARN }
                }
                'Orion.NPM.Interfaces' {
                    $srcIf = Get-SwisObject $Source ([string]$m.Definition)
                    $hit = $null
                    if ($srcIf) {
                        $hit = @(Invoke-Swql $Target 'SELECT Uri, InterfaceID, NodeID FROM Orion.NPM.Interfaces WHERE FullName = @fn' @{ fn = $srcIf.FullName }) | Select-Object -First 1
                    }
                    if ($hit) { $resolved.Add(@{ Name = "Orion.Nodes.NodeID=$($hit.NodeID) AND Interfaces.InterfaceID=$($hit.InterfaceID)"; Definition = [string]$hit.Uri }) }
                    else      { Write-Log "Group '$GroupName': interface $([string]$m.Name) not found on target, member skipped" -Level WARN }
                }
                'Orion.Groups' {
                    # Sub-groups are linked via the parent/child hierarchy when the child is created.
                }
                default {
                    Write-Log "Group '$GroupName': member type '$($m.Entity)' ($($m.Name)) is not migrated" -Level WARN
                }
            }
        }
        catch {
            Write-Log "Group '$GroupName': member '$($m.Name)' could not be resolved: $($_.Exception.Message)" -Level WARN
        }
    }

    # A group needs at least one member definition; v1.7 used this dummy filter for "empty" groups.
    if ($resolved.Count -eq 0) {
        $resolved.Add(@{ Name = 'Dummy Devices'; Definition = "filter:/Orion.Nodes[Vendor='Dummy']" })
    }
    return , $resolved.ToArray()
}

function New-TargetGroup {
    param($Target, $Group, $Members, $ParentId)

    $items = ($Members | ForEach-Object {
        '<MemberDefinitionInfo><Name>{0}</Name><Definition>{1}</Definition></MemberDefinitionInfo>' -f
            [System.Security.SecurityElement]::Escape([string]$_.Name),
            [System.Security.SecurityElement]::Escape([string]$_.Definition)
    }) -join ''
    $membersXml = ([xml]"<ArrayOfMemberDefinitionInfo xmlns='http://schemas.solarwinds.com/2008/Orion'>$items</ArrayOfMemberDefinitionInfo>").DocumentElement

    $owner = if (Test-IsEmpty $Group.Owner) { 'Core' } else { [string]$Group.Owner }

    # Arguments: [ParentID,] name, owner, refresh frequency, status rollup mode
    # (0 = mixed shows warning, 1 = worst, 2 = best), description, polling enabled, members
    if ($ParentId) {
        $verb = 'CreateContainerWithParent'
        $verbArgs = @($ParentId, $Group.Name, $owner, $Group.Frequency, $Group.StatusCalculator, $Group.Description, $Group.PollingEnabled, $membersXml)
    }
    else {
        $verb = 'CreateContainer'
        $verbArgs = @($Group.Name, $owner, $Group.Frequency, $Group.StatusCalculator, $Group.Description, $Group.PollingEnabled, $membersXml)
    }
    return (Invoke-SwisVerb $Target 'Orion.Container' $verb $verbArgs).InnerText
}

function Get-GroupDepth {
    param([int]$Id, [hashtable]$ParentOf)
    $depth = 0
    $seen  = @{}
    while ($ParentOf.ContainsKey($Id) -and -not $seen.ContainsKey($Id)) {   # $seen guards against cycles
        $seen[$Id] = $true
        $Id = $ParentOf[$Id]
        $depth++
    }
    return $depth
}

function Copy-Groups {
    [CmdletBinding(SupportsShouldProcess)]
    param($Source, $Target)

    # --- read all source groups ---
    $groups = [System.Collections.Generic.List[object]]::new()
    foreach ($uri in @(Invoke-Swql $Source 'SELECT Uri FROM Orion.Groups')) { $groups.Add((Get-SwisObject $Source $uri)) }
    Write-Log "Source has $($groups.Count) group(s)"

    $nameById = @{}
    foreach ($g in $groups) { $nameById[[int]$g.ContainerID] = [string]$g.Name }

    # --- child -> parent map. The child id is extracted exactly (v1.7 used LIKE '%ContainerID=5%',
    #     which also matched 50, 51, 5xx ...) ---
    $parentOf = @{}
    $links = @(Invoke-Swql $Source "SELECT ContainerID, Name, Definition FROM Orion.ContainerMemberDefinition WHERE Entity = 'Orion.Groups' OR Name LIKE '%Orion.Groups.ContainerID=%'")
    foreach ($link in $links) {
        foreach ($text in @([string]$link.Name, [string]$link.Definition)) {
            if ($text -match 'Orion\.Groups\.ContainerID\s*=\s*(\d+)(?!\d)' -or $text -match 'Groups/ContainerID\s*=\s*(\d+)(?!\d)') {
                $childId = [int]$Matches[1]
                if (-not $parentOf.ContainsKey($childId)) { $parentOf[$childId] = [int]$link.ContainerID }
                break
            }
        }
    }

    # --- parents first, so children can be attached to an existing parent ---
    $ordered = $groups | Sort-Object -Property @{ Expression = { Get-GroupDepth -Id ([int]$_.ContainerID) -ParentOf $parentOf } }, @{ Expression = { $_.Name } }

    foreach ($g in $ordered) {
        $name = [string]$g.Name
        try {
            if (Invoke-Swql $Target 'SELECT ContainerID FROM Orion.Container WHERE Name = @NewName' @{ NewName = $name }) {
                Write-Log "Group '$name' already exists on target, skipped"
                $script:Stats.GroupsSkipped++
                continue
            }

            $parentTargetId = $null
            if ($parentOf.ContainsKey([int]$g.ContainerID)) {
                $parentName = $nameById[$parentOf[[int]$g.ContainerID]]
                $parentTargetId = @(Invoke-Swql $Target 'SELECT ContainerID FROM Orion.Container WHERE Name = @n' @{ n = $parentName }) | Select-Object -First 1
                if (-not $parentTargetId -and $WhatIfPreference) {
                    Write-Log "Group '$name': parent '$parentName' would be created first (-WhatIf)" -Level DEBUG
                }
                elseif (-not $parentTargetId) {
                    throw "Parent group '$parentName' does not exist on the target (failed or skipped earlier). Fix that and re-run."
                }
            }

            $members = Resolve-ContainerMembers -Source $Source -Target $Target -GroupName $name -Members @(
                Invoke-Swql $Source 'SELECT Name, Definition, Entity FROM Orion.ContainerMemberDefinition WHERE ContainerID = @ID' @{ ID = $g.ContainerID }
            )

            if ($PSCmdlet.ShouldProcess($name, 'Create group on target')) {
                [void](New-TargetGroup -Target $Target -Group $g -Members $members -ParentId $parentTargetId)
                Write-Log "Created group '$name'$(if ($parentTargetId) { " below '$parentName'" })"
                $script:Stats.GroupsCreated++
            }
        }
        catch {
            Write-Log "Group '$name' FAILED: $($_.Exception.Message)" -Level ERROR
            Write-Log $_.ScriptStackTrace -Level DEBUG
            $script:Stats.GroupsFailed++
            Add-Problem 'GroupFailed' $name $_.Exception.Message
        }
    }
}

#endregion

#region Main ------------------------------------------------------------------------------------

Write-Log "Group migration started: $SourceHost -> $DestinationHost (WhatIf: $([bool]$WhatIfPreference))"
Write-Log "Log file: $LogPath"

try {
    $sourceSwis = Connect-Orion -HostName $SourceHost      -Credential $SourceCredential      -Trusted:$Trusted -Role 'source'
    $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'

    # Pre-flight: group members are resolved against the nodes on the target.
    $targetNodeCount = Invoke-Swql $targetSwis 'SELECT COUNT(NodeID) AS Cnt FROM Orion.Nodes'
    if (Test-IsEmpty $targetNodeCount) { $targetNodeCount = 0 }
    if ([int]$targetNodeCount -eq 0) {
        Write-Log 'The target has no nodes. Run MigrationScriptSolarwindsPlatforms-CPNodes-v2.0.ps1 first, otherwise all node members are skipped (and are not repaired by a re-run).' -Level WARN
    }
    else {
        Write-Log "Target has $targetNodeCount node(s)"
    }

    Copy-Groups -Source $sourceSwis -Target $targetSwis
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
