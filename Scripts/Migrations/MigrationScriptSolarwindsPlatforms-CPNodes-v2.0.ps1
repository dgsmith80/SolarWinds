#Requires -Version 5.1
#Requires -Modules SwisPowerShell
<#
.SYNOPSIS
    Step 1 of 2: migrates custom property definitions, nodes, interfaces, volumes, pollers and NCM
    membership from one SolarWinds Platform (Orion) instance to another via SWIS.
    Groups are migrated by the separate script MigrationScriptSolarwindsPlatforms-Groups-v2.0.ps1,
    which must be run AFTER this one (group members are resolved against the nodes on the target).
    This script runs standalone.

.DESCRIPTION
    Author:     Holger Mundt
    Email:      holger.mundt@loop1.com
    Version:    2.0
    History:    2023-09 HM - Creation of Script
                2023-10 HM - Extension and Bugfixing
                v1.7    HM - (last version before refactoring)
                v2.0    Refactoring of v1.7:
                          - one code path for SNMP v1/v2, SNMPv3 and WMI nodes
                          - table-driven custom property sync (fixes IPAM CP being created as interface CP)
                          - volume custom properties are now actually copied (variable name bug)
                          - per-node try/catch with rollback, so a re-run can retry cleanly
                          - -WhatIf support, log file, summary, exit code
                          - no hard-coded credentials

    Split from MigrationScriptSolarwindsPlatforms-v2.0.ps1: this file contains everything except groups.

    What is NOT migrated (same as v1.7): Agent nodes, groups (see the Groups script),
    values of group/application/alert/report/IPAM custom properties, WMI credentials themselves
    (they must exist on the target with the same name).

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

.PARAMETER EngineId
    Pin all migrated nodes to this polling engine. Default (0): spread nodes over all engines of
    the target, always picking the engine with the fewest nodes.

.PARAMETER FallbackWmiCredentialName
    Name of a credential on the target to use when the source WMI credential name does not exist
    there. Default: none - such nodes fail and are listed in the summary.

.PARAMETER KeepPartialNodes
    Keep nodes whose interfaces/volumes/pollers/custom properties could not be fully copied.
    Default: such nodes are removed from the target again so that a re-run can retry them.

.PARAMETER ThrottleMs
    Pause (ms) before each node that is created. Default 500.

.PARAMETER LogPath
    Log file. Default: next to the script, with timestamp.

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-CPNodes-v2.0.ps1 -SourceHost old.contoso.local -DestinationHost new.contoso.local -WhatIf

.EXAMPLE
    .\MigrationScriptSolarwindsPlatforms-CPNodes-v2.0.ps1 -SourceHost 192.168.50.106 -DestinationHost 192.168.2.200 -Verbose
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourceHost,
    [Parameter(Mandatory)][string]$DestinationHost,
    [pscredential]$SourceCredential,
    [pscredential]$DestinationCredential,
    [switch]$Trusted,
    [int]$EngineId = 0,
    [string]$FallbackWmiCredentialName,
    [switch]$KeepPartialNodes,
    [int]$ThrottleMs = 500,
    [string]$LogPath = (Join-Path -Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) `
                                 -ChildPath ('SwisMigration_CpNodes_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date)))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Configuration ---------------------------------------------------------------------------

# Properties copied for nodes / interfaces / volumes. The SELECT statements are built from these lists.
$NodePropsToCopy = @(
    'AgentPort', 'Allow64BitCounters', 'BufferBgMissThisHour', 'BufferBgMissToday', 'BufferHgMissThisHour', 'BufferHgMissToday',
    'BufferLgMissThisHour', 'BufferLgMissToday', 'BufferMdMissThisHour', 'BufferMdMissToday', 'BufferNoMemThisHour', 'BufferNoMemToday',
    'BufferSmMissThisHour', 'BufferSmMissToday', 'Caption', 'ChildStatus', 'CMTS', 'Community', 'Contact', 'DNS', 'DynamicIP', 'External',
    'EntityType', 'GroupStatus', 'IOSImage', 'IOSVersion', 'IPAddress', 'IPAddressGUID', 'IPAddressType', 'LastSystemUpTimePollUtc',
    'Location', 'MachineType', 'NodeDescription', 'ObjectSubType', 'PollInterval', 'RediscoveryInterval', 'Severity', 'SNMPVersion',
    'StatCollection', 'Status', 'StatusDescription', 'StatusLED', 'SysName', 'SysObjectID', 'TotalMemory', 'Vendor', 'VendorIcon',
    'RWCommunity'
)

$InterfacePropsToCopy = @(
    'AdminStatus', 'AdminStatusLED', 'Caption', 'Counter64', 'CustomBandwidth', 'FullName',
    'IfName', 'InBandwidth', 'Inbps', 'InDiscardsThisHour', 'InDiscardsToday', 'InErrorsThisHour', 'InErrorsToday',
    'InMcastPps', 'InPercentUtil', 'InPktSize', 'InPps', 'InterfaceAlias', 'InterfaceIcon', 'InterfaceIndex',
    'InterfaceMTU', 'InterfaceName', 'InterfaceSpeed', 'InterfaceSubType', 'InterfaceType', 'InterfaceTypeDescription',
    'InterfaceTypeName', 'InUcastPps', 'MaxInBpsToday', 'MaxOutBpsToday', 'ObjectSubType', 'OperStatus', 'OutBandwidth',
    'Outbps', 'OutDiscardsThisHour', 'OutDiscardsToday', 'OutErrorsThisHour', 'OutErrorsToday', 'OutMcastPps',
    'OutPercentUtil', 'OutPktSize', 'OutPps', 'OutUcastPps', 'PhysicalAddress', 'PollInterval', 'RediscoveryInterval',
    'Severity', 'StatCollection', 'Status', 'StatusLED', 'UnManaged', 'UnPluggable'
)

$VolumePropsToCopy = @(
    'Icon', 'Index', 'Caption', 'StatusIcon', 'Type', 'Size', 'Responding', 'FullName',
    'VolumePercentUsed', 'VolumeAllocationFailuresThisHour', 'VolumeDescription', 'VolumeSpaceUsed',
    'VolumeAllocationFailuresToday', 'VolumeSpaceAvailable', 'VolumeTypeID'
)

$NodeSelect = 'SELECT ' + ($NodePropsToCopy -join ', ') + ' FROM Orion.Nodes WHERE NodeID = @nodeid'

# Custom property tables: Orion.CustomProperty.Table -> entity used for CreateCustomProperty on the target.
#   Padding = the empty trailing verb arguments after Field, Description, DataType, MaxLength
#   FixNegativeLength = some built-in properties report MaxLength -1 -> set to 0
$CpMap = @{
    AlertConfigurationsCustomProperties = @{ Entity = 'Orion.AlertConfigurationsCustomProperties'; Padding = @('', '', '', '', '', '') }
    APM_ApplicationCustomProperties     = @{ Entity = 'Orion.APM.ApplicationCustomProperties';     Padding = @('', '', '', '', '', '', '', $false, '', '', '') }
    ContainerCustomProperties           = @{ Entity = 'Orion.GroupCustomProperties';               Padding = @('', '', '', '', '', '') }
    Interfaces                          = @{ Entity = 'Orion.NPM.InterfacesCustomProperties';      Padding = @('', '', '', '', '', '') }
    NodesCustomProperties               = @{ Entity = 'Orion.NodesCustomProperties';               Padding = @('', '', '', '', '', ''); FixNegativeLength = $true }
    ReportDefinitions                   = @{ Entity = 'Orion.ReportsCustomProperties';             Padding = @('', '', '', '', '', '') }
    Volumes                             = @{ Entity = 'Orion.VolumesCustomProperties';             Padding = @('', '', '', '', '', '') }
    # v1.7 created these as *interface* custom properties (copy/paste bug). Not supported via this verb.
    IPAM_GroupAttrData                  = @{ Unsupported = 'IPAM custom properties must be created in IPAM on the target manually' }
}

# Node batches: SNMP/ICMP v1/v2, SNMPv3, WMI. Agent nodes are not migrated.
$NodeSets = @(
    @{ Label = 'SNMP/ICMP (v1, v2, none)'; Mode = 'Snmp';   Where = "ObjectSubType NOT IN ('WMI', 'Agent') AND SNMPVersion IN ('2','1','0')" }
    @{ Label = 'SNMPv3';                   Mode = 'SnmpV3'; Where = "ObjectSubType NOT IN ('WMI', 'Agent') AND SNMPVersion = '3'" }
    @{ Label = 'WMI';                      Mode = 'Wmi';    Where = "ObjectSubType = 'WMI'" }
)

#endregion

#region State -----------------------------------------------------------------------------------

$script:Stats = [ordered]@{
    CustomPropertiesCreated = 0
    NodesCreated = 0; NodesPartial = 0; NodesSkipped = 0; NodesFailed = 0
}
$script:Problems   = [System.Collections.Generic.List[object]]::new()
$script:CpFields   = @{}   # Table -> List[string] of custom property names on the source
$script:TargetIps  = $null # HashSet[string] of IPs already on the target
$script:EngineLoad = @{}   # EngineID -> node count on the target
$script:NcmNodes   = $null # HashSet[int] of source CoreNodeIDs that are in NCM

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

function Get-CustomPropertyValues {
    # Returns a hashtable Name -> Value of the given custom properties of one object.
    # The key column is selected as well: with a single custom property Get-SwisData would otherwise
    # return a bare scalar instead of an object. Null values are skipped.
    param($Swis, [string]$Entity, [string]$KeyName, $KeyValue, [string[]]$Fields)
    $values = @{}
    if (-not $Fields -or @($Fields).Count -eq 0) { return $values }
    $columns = (@($KeyName) + @($Fields)) -join ', '
    $row = @(Invoke-Swql $Swis "SELECT $columns FROM $Entity WHERE $KeyName = @key" @{ key = $KeyValue }) | Select-Object -First 1
    if ($null -eq $row) { return $values }
    foreach ($prop in $row.psobject.Properties) {
        if ($prop.Name -eq $KeyName) { continue }
        if ($null -eq $prop.Value -or $prop.Value -is [DBNull]) { continue }
        $values[$prop.Name] = $prop.Value
    }
    return $values
}

function Get-CpFieldList {
    # Custom property names of one table as string[] (empty array if the source has none).
    param([string]$Table)
    if ($script:CpFields.ContainsKey($Table)) { return , $script:CpFields[$Table].ToArray() }
    return , @()
}

function Get-ReverseDnsName {
    # Reverse (PTR) lookup of an IP address. Returns '' if there is no DNS name.
    # (v1.7 wrote the literal string 'DNS' for WMI nodes without a DNS name.)
    param([string]$IPAddress)
    if (Test-IsEmpty $IPAddress) { return '' }
    try {
        $name = [System.Net.Dns]::GetHostEntry($IPAddress).HostName
        # some resolvers echo the IP back when there is no PTR record
        if ((Test-IsEmpty $name) -or $name -eq $IPAddress) { return '' }
        return [string]$name
    }
    catch {
        Write-Log "Reverse DNS lookup for $IPAddress returned nothing: $($_.Exception.Message)" -Level DEBUG
        return ''
    }
}

function Test-CustomPropertyExists {
    param($Swis, [string]$Table, [string]$Field)
    $hit = Invoke-Swql $Swis 'SELECT Field FROM Orion.CustomProperty WHERE Field = @FieldValue AND Table = @Table' `
                       @{ FieldValue = $Field; Table = $Table }
    return [bool]$hit
}

function Sync-CustomProperties {
    # Collects the custom property names per table from the source and creates missing ones on the target.
    [CmdletBinding(SupportsShouldProcess)]
    param($Source, $Target)

    $all = @(Invoke-Swql $Source 'SELECT Table, Field, DataType, MaxLength, Description FROM Orion.CustomProperty')
    Write-Log "Source has $($all.Count) custom properties"

    foreach ($cp in $all) {
        $table = [string]$cp.Table
        $field = [string]$cp.Field

        if (-not $script:CpFields.ContainsKey($table)) { $script:CpFields[$table] = [System.Collections.Generic.List[string]]::new() }
        $script:CpFields[$table].Add($field)

        $map = $CpMap[$table]
        if (-not $map) {
            Write-Log "Custom property ${table}.${field}: table not supported by this script, not created" -Level WARN
            continue
        }
        if ($map.ContainsKey('Unsupported')) {
            Write-Log "Custom property $table.${field}: $($map.Unsupported)" -Level WARN
            continue
        }
        if (Test-CustomPropertyExists -Swis $Target -Table $table -Field $field) {
            Write-Log "Custom property ${table}.${field} already exists" -Level DEBUG
            continue
        }

        $length = $cp.MaxLength
        if ($map.ContainsKey('FixNegativeLength') -and $length -eq -1) { $length = '0' }
        $verbArgs = @($field, $cp.Description, $cp.DataType, $length) + $map.Padding

        if ($PSCmdlet.ShouldProcess("${table}.${field}", 'Create custom property on target')) {
            try {
                Invoke-SwisVerb $Target $map.Entity 'CreateCustomProperty' $verbArgs | Out-Null
                $script:Stats.CustomPropertiesCreated++
                Write-Log "Created custom property ${table}.${field}"
            }
            catch {
                Write-Log "Creating custom property ${table}.${field} failed: $($_.Exception.Message)" -Level ERROR
                Add-Problem 'CustomProperty' "${table}.${field}" $_.Exception.Message
            }
        }
    }

    # Marker property for migrated nodes
    if (-not (Test-CustomPropertyExists -Swis $Target -Table 'NodesCustomProperties' -Field 'MigratedByAPI')) {
        if ($PSCmdlet.ShouldProcess('NodesCustomProperties.MigratedByAPI', 'Create marker custom property on target')) {
            Invoke-SwisVerb $Target 'Orion.NodesCustomProperties' 'CreateCustomProperty' `
                @('MigratedByAPI', 'CopiedFromOldEnvironmentbyScript', 'System.String', 100, '', '', '', '', '', '') | Out-Null
        }
    }
}

#endregion

#region Engines and NCM ------------------------------------------------------------------------

function Initialize-EngineLoad {
    param($Target)
    $script:EngineLoad = @{}
    foreach ($id in @(Invoke-Swql $Target 'SELECT EngineID FROM Orion.Engines')) { $script:EngineLoad[[int]$id] = 0 }
    foreach ($row in @(Invoke-Swql $Target 'SELECT EngineID, COUNT(NodeID) AS Cnt FROM Orion.Nodes GROUP BY EngineID')) {
        $script:EngineLoad[[int]$row.EngineID] = [int]$row.Cnt
    }
    if ($script:EngineLoad.Count -eq 0) { throw 'No polling engine found on the target.' }
}

function Get-TargetEngineId {
    # Engine with the fewest nodes (including engines that have none yet), or the pinned engine.
    if ($EngineId -gt 0) { return $EngineId }
    $pick = $script:EngineLoad.GetEnumerator() | Sort-Object -Property Value, Key | Select-Object -First 1
    return [int]$pick.Key
}

function Initialize-NcmState {
    param($Source, $Target)
    $script:NcmNodes = [System.Collections.Generic.HashSet[int]]::new()
    try {
        foreach ($id in @(Invoke-Swql $Source 'SELECT CoreNodeID FROM NCM.Nodes')) {
            if (-not (Test-IsEmpty $id)) { [void]$script:NcmNodes.Add([int]$id) }
        }
        [void](Invoke-Swql $Target 'SELECT TOP 1 NodeID FROM Cirrus.Nodes')   # NCM available on the target?
        Write-Log "NCM: $($script:NcmNodes.Count) source node(s) are in NCM"
    }
    catch {
        Write-Log "NCM not available on source and/or target - NCM membership will not be migrated ($($_.Exception.Message))" -Level WARN
        $script:NcmNodes.Clear()
    }
}

#endregion

#region Node copy -------------------------------------------------------------------------------

function Copy-Pollers {
    param($Source, $Target, [string]$Prefix, $SourceId, $NewId, [System.Collections.Generic.List[string]]$Issues)
    $pollerTypes = @(Invoke-Swql $Source 'SELECT PollerType FROM Orion.Pollers WHERE NetObject = @netobject' `
                                 @{ netobject = "${Prefix}:$SourceId" })
    foreach ($pollerType in $pollerTypes) {
        try {
            Write-Log "    Adding poller $pollerType" -Level DEBUG
            New-SwisObject $Target -EntityType 'Orion.Pollers' -Properties @{
                PollerType    = "$pollerType"
                NetObject     = "${Prefix}:$NewId"
                NetObjectType = $Prefix
                NetObjectID   = $NewId
            } | Out-Null
        }
        catch {
            $Issues.Add("Poller $pollerType on ${Prefix}:$NewId - $($_.Exception.Message)")
        }
    }
}

function Copy-NodeInterfaces {
    param($Source, $Target, $SourceNode, $NewNodeId, [System.Collections.Generic.List[string]]$Issues)
    $rows = @(Invoke-Swql $Source 'SELECT Nodes.Interfaces.Uri, Nodes.Interfaces.InterfaceID FROM Orion.Nodes WHERE NodeID = @node' `
                          @{ node = $SourceNode.NodeID })
    foreach ($row in $rows) {
        if (Test-IsEmpty $row.Uri) { continue }   # node without interfaces returns one empty row
        try {
            $src = Get-SwisObject $Source $row.Uri
            if (-not $src -or $src.HasObsoleteData -eq 1) { continue }

            Write-Log "  Copying interface $($SourceNode.Caption) / $($src.Caption)" -Level DEBUG
            $props = @{}
            foreach ($p in $InterfacePropsToCopy) { $props[$p] = $src[$p] }
            $props['NodeID'] = $NewNodeId

            $newUri = New-SwisObject $Target -EntityType 'Orion.NPM.Interfaces' -Properties $props
            $newIf  = Get-SwisObject $Target $newUri

            $cps = Get-CustomPropertyValues -Swis $Source -Entity 'Orion.NPM.InterfacesCustomProperties' -KeyName 'InterfaceID' `
                       -KeyValue $row.InterfaceID -Fields (Get-CpFieldList 'Interfaces')
            if ($cps.Count -gt 0) { Set-SwisObject $Target ($newUri + '/CustomProperties') $cps }

            Copy-Pollers -Source $Source -Target $Target -Prefix 'I' -SourceId $row.InterfaceID -NewId $newIf.InterfaceID -Issues $Issues
        }
        catch {
            $Issues.Add("Interface $($row.InterfaceID) - $($_.Exception.Message)")
        }
    }
}

function Copy-NodeVolumes {
    param($Source, $Target, $SourceNode, $NewNodeId, [System.Collections.Generic.List[string]]$Issues)
    $rows = @(Invoke-Swql $Source 'SELECT Nodes.Volumes.Uri, Nodes.Volumes.VolumeID FROM Orion.Nodes WHERE NodeID = @node' `
                          @{ node = $SourceNode.NodeID })
    foreach ($row in $rows) {
        if (Test-IsEmpty $row.VolumeID) { continue }
        try {
            $src = Get-SwisObject $Source $row.Uri
            Write-Log "  Copying volume $($SourceNode.Caption) / $($src.Caption)" -Level DEBUG
            $props = @{}
            foreach ($p in $VolumePropsToCopy) { $props[$p] = $src[$p] }
            $props['NodeID'] = $NewNodeId

            $newUri = New-SwisObject $Target -EntityType 'Orion.Volumes' -Properties $props
            $newVol = Get-SwisObject $Target $newUri

            # v1.7 checked a non-existing variable here, so volume custom properties were never copied.
            $cps = Get-CustomPropertyValues -Swis $Source -Entity 'Orion.VolumesCustomProperties' -KeyName 'VolumeID' `
                       -KeyValue $row.VolumeID -Fields (Get-CpFieldList 'Volumes')
            if ($cps.Count -gt 0) { Set-SwisObject $Target ($newUri + '/CustomProperties') $cps }

            Copy-Pollers -Source $Source -Target $Target -Prefix 'V' -SourceId $row.VolumeID -NewId $newVol.VolumeID -Issues $Issues
        }
        catch {
            $Issues.Add("Volume $($row.VolumeID) - $($_.Exception.Message)")
        }
    }
}

function Get-TargetWmiCredentialId {
    # Maps the WMI credential of a source node to the credential with the same name on the target.
    # No silent "first credential with a backslash" fallback (v1.7): a wrong credential can lock accounts.
    param($Source, $Target, $SourceNodeId)

    $credId = Invoke-Swql $Source "SELECT SettingValue FROM Orion.NodeSettings WHERE SettingName = 'WMICredential' AND NodeID = @NodeID" `
                          @{ NodeID = $SourceNodeId }
    $credName = $null
    if (-not (Test-IsEmpty $credId)) {
        $credName = Invoke-Swql $Source 'SELECT Name FROM Orion.Credential WHERE ID = @ID' @{ ID = [int]@($credId)[0] }
    }

    foreach ($name in @($credName, $FallbackWmiCredentialName)) {
        if (Test-IsEmpty $name) { continue }
        $hit = @(Invoke-Swql $Target "SELECT ID FROM Orion.Credential WHERE CredentialOwner = 'Orion' AND Name = @Name" @{ Name = [string]@($name)[0] })
        if ($hit.Count -gt 1) { Write-Log "More than one credential named '$name' on target - using the first" -Level WARN }
        if ($hit.Count -ge 1) { return $hit[0] }
    }
    return $null
}

function Copy-SwisNode {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        $Source, $Target, $SourceNode,
        [Parameter(Mandatory)][ValidateSet('Snmp', 'SnmpV3', 'Wmi')][string]$Mode
    )

    $ip    = [string]$SourceNode.IPAddress
    $label = "$($SourceNode.Caption) ($ip)"

    if ($script:TargetIps.Contains($ip)) {
        Write-Log "Skipping $label - IP already exists on $DestinationHost"
        $script:Stats.NodesSkipped++
        return
    }
    if (-not $PSCmdlet.ShouldProcess($label, "Create $Mode node on target")) { return }

    Start-Sleep -Milliseconds $ThrottleMs
    Write-Log "Copying $label [$Mode]"

    $issues  = [System.Collections.Generic.List[string]]::new()
    $failure = $null
    $newUri  = $null
    $engine  = $null

    try {
        $src = @(Invoke-Swql $Source $NodeSelect @{ nodeid = $SourceNode.NodeID }) | Select-Object -First 1
        if (-not $src) { throw 'Source node properties could not be read.' }

        $props = @{}
        foreach ($p in $NodePropsToCopy) { $props[$p] = $src.$p }
        $engine = Get-TargetEngineId
        $props['EngineID'] = $engine

        switch ($Mode) {
            'SnmpV3' {
                $v3 = @(Invoke-Swql $Source 'SELECT Username, PrivacyMethod, PrivacyKey, PrivacyKeyIsPassword, AuthenticationMethod, AuthenticationKey, AuthenticationKeyIsPassword FROM Orion.SNMPv3Credentials WHERE NodeID = @NodeID' `
                                    @{ NodeID = $SourceNode.NodeID }) | Select-Object -First 1
                if (-not $v3) { throw 'No SNMPv3 credentials found on the source node.' }
                $props['SNMPV3Username']      = $v3.Username
                $props['SNMPV3PrivMethod']    = $v3.PrivacyMethod
                $props['SNMPV3PrivKey']       = $v3.PrivacyKey
                $props['SNMPV3PrivKeyIsPwd']  = $v3.PrivacyKeyIsPassword
                $props['SNMPV3AuthMethod']    = $v3.AuthenticationMethod
                $props['SNMPV3AuthKey']       = $v3.AuthenticationKey
                $props['SNMPV3AuthKeyIsPwd']  = $v3.AuthenticationKeyIsPassword
            }
            'Wmi' {
                # WMI nodes need a DNS value: use the source value, else a reverse lookup of the IP, else empty
                if (Test-IsEmpty $props['DNS']) { $props['DNS'] = Get-ReverseDnsName -IPAddress $ip }
            }
        }

        # --- create the node ---
        $newUri = New-SwisObject $Target -EntityType 'Orion.Nodes' -Properties $props
        [void]$script:TargetIps.Add($ip)
        if (-not $script:EngineLoad.ContainsKey([int]$engine)) { $script:EngineLoad[[int]$engine] = 0 }
        $script:EngineLoad[[int]$engine]++
        $newNode   = Get-SwisObject $Target $newUri
        $newNodeId = $newNode.NodeID

        # --- custom properties (marker first, so it survives a failure of the copied values) ---
        try { Set-SwisObject $Target ($newUri + '/CustomProperties') @{ MigratedByAPI = 'true' } }
        catch { $issues.Add("MigratedByAPI marker - $($_.Exception.Message)") }

        try {
            $cps = Get-CustomPropertyValues -Swis $Source -Entity 'Orion.NodesCustomProperties' -KeyName 'NodeID' `
                       -KeyValue $SourceNode.NodeID -Fields (Get-CpFieldList 'NodesCustomProperties')
            if ($cps.Count -gt 0) { Set-SwisObject $Target ($newUri + '/CustomProperties') $cps }
        }
        catch { $issues.Add("Custom properties - $($_.Exception.Message)") }

        # --- WMI credential ---
        if ($Mode -eq 'Wmi') {
            $wmiCredId = Get-TargetWmiCredentialId -Source $Source -Target $Target -SourceNodeId $SourceNode.NodeID
            if ($null -eq $wmiCredId) {
                $issues.Add('No matching WMI credential on target (create it with the same name, or use -FallbackWmiCredentialName)')
            }
            else {
                try {
                    New-SwisObject $Target -EntityType 'Orion.NodeSettings' -Properties @{
                        NodeID = $newNodeId; SettingName = 'WMICredential'; SettingValue = $wmiCredId
                    } | Out-Null
                }
                catch { $issues.Add("WMI credential assignment - $($_.Exception.Message)") }
            }
        }

        # --- pollers, interfaces, volumes, NCM ---
        Copy-Pollers        -Source $Source -Target $Target -Prefix 'N' -SourceId $SourceNode.NodeID -NewId $newNodeId -Issues $issues
        Copy-NodeInterfaces -Source $Source -Target $Target -SourceNode $SourceNode -NewNodeId $newNodeId -Issues $issues
        Copy-NodeVolumes    -Source $Source -Target $Target -SourceNode $SourceNode -NewNodeId $newNodeId -Issues $issues

        if ($script:NcmNodes.Contains([int]$SourceNode.NodeID)) {
            try { Invoke-SwisVerb $Target 'Cirrus.Nodes' 'AddNodeToNCM' @($newNodeId) | Out-Null }
            catch { $issues.Add("Add to NCM - $($_.Exception.Message)") }
        }
    }
    catch {
        $failure = $_.Exception.Message
    }

    # --- result handling ---
    if (-not $failure -and $issues.Count -eq 0) {
        $script:Stats.NodesCreated++
        return
    }

    $detail = (@($failure) + $issues.ToArray() | Where-Object { $_ }) -join ' | '

    if ($newUri -and $KeepPartialNodes) {
        Write-Log "Node $label kept but INCOMPLETE: $detail" -Level WARN
        $script:Stats.NodesPartial++
        Add-Problem 'NodePartial' $label $detail
        return
    }

    if ($newUri) {
        try {
            Remove-SwisObject $Target $newUri
            [void]$script:TargetIps.Remove($ip)
            $script:EngineLoad[[int]$engine] = $script:EngineLoad[[int]$engine] - 1
            $detail += ' | node removed from target again (re-run will retry it)'
        }
        catch {
            $detail += " | ROLLBACK FAILED, remove manually: $newUri ($($_.Exception.Message))"
        }
    }
    Write-Log "Node $label FAILED: $detail" -Level ERROR
    $script:Stats.NodesFailed++
    Add-Problem 'NodeFailed' $label $detail
}

#endregion

#region Main ------------------------------------------------------------------------------------

Write-Log "CP + Nodes migration started: $SourceHost -> $DestinationHost (WhatIf: $([bool]$WhatIfPreference), KeepPartialNodes: $([bool]$KeepPartialNodes))"
Write-Log "Log file: $LogPath"

try {
    $sourceSwis = Connect-Orion -HostName $SourceHost      -Credential $SourceCredential      -Trusted:$Trusted -Role 'source'
    $targetSwis = Connect-Orion -HostName $DestinationHost -Credential $DestinationCredential -Trusted:$Trusted -Role 'destination'

    # 1. custom property definitions
    Sync-CustomProperties -Source $sourceSwis -Target $targetSwis

    # 2. lookup state, loaded once instead of one query per node
    $script:TargetIps = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($existingIp in @(Invoke-Swql $targetSwis 'SELECT IPAddress FROM Orion.Nodes')) { [void]$script:TargetIps.Add([string]$existingIp) }
    Initialize-EngineLoad -Target $targetSwis
    Initialize-NcmState   -Source $sourceSwis -Target $targetSwis

    $agentCount = Invoke-Swql $sourceSwis "SELECT COUNT(NodeID) AS Cnt FROM Orion.Nodes WHERE ObjectSubType = 'Agent'"
    if ($agentCount -and [int]$agentCount -gt 0) {
        Write-Log "$agentCount agent node(s) on the source are NOT migrated by this script" -Level WARN
    }

    # 3. nodes
    foreach ($set in $NodeSets) {
        $nodes = @(Invoke-Swql $sourceSwis "SELECT Uri, IPAddress, Caption, NodeID, ObjectSubType FROM Orion.Nodes WHERE $($set.Where)")
        Write-Log "$($set.Label): $($nodes.Count) node(s) on source"
        foreach ($node in $nodes) {
            try {
                Copy-SwisNode -Source $sourceSwis -Target $targetSwis -SourceNode $node -Mode $set.Mode
            }
            catch {
                Write-Log "Node $($node.Caption) ($($node.IPAddress)) FAILED unexpectedly: $($_.Exception.Message)" -Level ERROR
                $script:Stats.NodesFailed++
                Add-Problem 'NodeFailed' "$($node.Caption) ($($node.IPAddress))" $_.Exception.Message
            }
        }
    }

    Write-Log 'Next step: run MigrationScriptSolarwindsPlatforms-Groups-v2.0.ps1 to migrate the groups.'
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
