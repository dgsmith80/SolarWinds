# Title: SNMPv2Tov3
# By: Philippe Langlois
# Version: 1.1
#Config variables
$All_Nodes = @( 'Node1', 'Node2', 'Node3' ) #Use captions (or adapt script below if you use IP addresses)
$ROSNMPCredentialID = 31
$RWSNMPCredentialID = $null
# To Get Creds IDs: 
# Get-SwisData -SwisConnection $swis "SELECT ID, Name FROM Orion.Credential where CredentialOwner='Orion' AND CredentialType like '%SnmpCredentialsV3'"
$base = "C:\Scripts\SolarWinds\test"
$LogPath = "$base\Logs"
$module_path = "C:\BMR\Scripts\PS_Modules"
[string]$LogfileName = "GroupMembership" #FileName of the Logfile
[int]$LogDeleteAfterDays = 10 #Time Period in Days when older Files will be deleted
$OrionServer = '1.2.3.4'
$VerbosePreference = "Continue"
#Start Logging:
[string]$LogFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
Start-Transcript -Path $logFile -Append -IncludeInvocationHeader 
#Import required modules:
Import-Module -DisableNameChecking $module_path"\PowerOrion\PowerOrion.psm1"
if (Get-Module -ListAvailable -Name SwisPowerShell) {
    Write-Host "[INFO] SwisPowershell is available."
} 
else {
    Write-Host "[ERROR] SwisPowershellmodule does not exist" -BackgroundColor Red
    Import-Module SwisPowershell
}
#connect to Orion API
IF ($null -eq $Swis) {
    $swis = Connect-Swis -Trusted -Hostname $OrionServer  #temporarily as per https://thwack.solarwinds.com/product-forums/the-orion-platform/f/orion-sdk/96684/unable-to-use--trusted-or--certificate-to-connect-to-swis-on-main-polling-engine    
}
try {
    Write-Host "Verifying connection to Orion."
    $swis.Open()
}
catch {
    Write-Warning 'The following message requires your attention:'
    Write-Host $_
    Write-Host $_.Exception.Message
}
#starting the show:
$NodesInfo = @() #Init
foreach ($node in $All_Nodes) {
    #getting initial Node information
    $temp = $null #init
    $temp = Get-SwisData -SwisConnection $swis @"
        SELECT TOP 1000 NodeID, EngineID, ObjectSubType, IPAddress, Caption, DNS, Description, SysName, Vendor, SysObjectID, Location AS SNMP_Location, Contact AS SNMP_Contact, SNMPVersion AS SNMP_Version, Community AS SNMP_Community, RWCommunity AS SNMP_RWCommunity, Status, CustomStatus, PolledStatus, StatusDescription, AgentPort, External, OrionIdPrefix, OrionIdColumn, EntityType, DetailsUrl, Uri
        FROM Orion.Nodes WHERE Caption = @C
"@ -Parameters @{ C = $node }
    $NodesInfo += $temp
}
#Adding Columns:
if (((($NodesInfo | Get-Member) | Where-Object { $_.name -like 'ChangeVersionCheck' }).count) -eq 0) {
    $NodesInfo | Add-Member -MemberType NoteProperty "ChangeVersionCheck" -Value $null  #adding a column
}
if (((($NodesInfo | Get-Member) | Where-Object { $_.name -like 'ROSNMPCredentialID' }).count) -eq 0) {
    $NodesInfo | Add-Member -MemberType NoteProperty "ROSNMPCredentialID" -Value $null  #adding a column
}
if (((($NodesInfo | Get-Member) | Where-Object { $_.name -like 'RWSNMPCredentialID' }).count) -eq 0) {
    $NodesInfo | Add-Member -MemberType NoteProperty "RWSNMPCredentialID" -Value $null  #adding a column
}
#Obtain the Credential set ID:
foreach ($node in $NodesInfo) {
    if ($node.SNMP_Version -eq 3) {
        $node.ROSNMPCredentialID = Get-SwisData $swis @"
        SELECT TOP 1 SettingValue FROM Orion.NodeSettings WHERE NodeID = @NodeID and SettingName = 'ROSNMPCredentialID'
"@ -Parameters @{ NodeID = $($node.NodeID) }
        $node.RWSNMPCredentialID = Get-SwisData $swis @"
        SELECT TOP 1 SettingValue FROM Orion.NodeSettings WHERE NodeID = @NodeID and SettingName = 'RWSNMPCredentialID'
"@ -Parameters @{ NodeID = $node.NodeID }   
    }
}
#validate each node SNMP Version
$snmpv3Properties = @{ SNMPVersion = 3 ; }
foreach ($node in $NodesInfo) {
    if ($node.SNMP_Version -in 1, 2) {
        $node.ChangeVersionCheck = $true
        Set-SwisObject -SwisConnection $Swis -Uri $node.Uri -Properties $snmpv3Properties #changing to v3
        #update credential set:
        if ($ROSNMPCredentialID) {
            Invoke-SwisVerb $swis 'Orion.Reporting' 'ExecuteSQL' "INSERT INTO NodeSettings (NodeID, SettingName, SettingValue) VALUES ($($node.NodeID), 'ROSNMPCredentialID', $ROSNMPCredentialID)"
        }
    }
}        
#Validate SNMPv3 Cresds after change:
foreach ($node in $NodesInfo) {
    if ($node.SNMP_Version -eq 3) {
        $node.ROSNMPCredentialID = Get-SwisData $swis @"
        SELECT TOP 1 SettingValue FROM Orion.NodeSettings WHERE NodeID = @NodeID and SettingName = 'ROSNMPCredentialID'
"@ -Parameters @{ NodeID = $($node.NodeID) }
        $node.RWSNMPCredentialID = Get-SwisData $swis @"
        SELECT TOP 1 SettingValue FROM Orion.NodeSettings WHERE NodeID = @NodeID and SettingName = 'RWSNMPCredentialID'
"@ -Parameters @{ NodeID = $node.NodeID }   
    }
}
foreach ($node in $NodesInfo) {
    if ($node.ChangeVersionCheck -eq $true) {
        if ($node.ROSNMPCredentialID -ne $ROSNMPCredentialID -or $node.RWSNMPCredentialID -ne $RWSNMPCredentialID) {
            Write-Host "Warning: $($node.Caption) does not have the expected SNMPv3 credential set."
        }
    }
}
Stop-Transcript