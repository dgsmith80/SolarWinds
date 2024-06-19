<##################################################################################################################################

Script:              Hardware Health Sensor Thresholds (Update thresholds and remove problematic sensors)

Author:              David Smith
Date created:        11.03.2022
Date Updated:        14.06.2024
Purpose:             Export SolarWinds Platform Custom Properties

Prerequisites:       SwisPowerShell (SolarWinds Orion Module) https://www.powershellgallery.com/packages/SwisPowerShell

##################################################################################################################################>

### Start Debug Transcript (Remove ## to enable)
#Start-Transcript -Path "D:\SolarWinds\Automation\Logging\HWSensorThresholds_Transcript.log" -Append -Force -NoClobber

## File/Registry Locations
$Log = 'D:\SolarWinds\Automation\Logging\HWSensorThresholds.log'
$RegistryKey = 'HKLM:\SOFTWARE\WOW6432Node\NWG_Solarwinds\'

## Load Modules
Import-Module SwisPowerShell  ## SwisPowerShell Module (Used to execute SWQL Commands)
Import-Module "D:\SolarWinds\Automation\PSModules\NWG-Orion.psm1" -Force  ## NWG Orion Module (Common Function Library)

### Server Details
## SWIS Connection Info
$OrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MLV
$SecondaryOrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MAC
$SQLServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-SQL-Orion

### SWIS Connection Test
$SWISConnectionTest = "SELECT Name, InstanceType, InstanceSiteID FROM Orion.SwisFeature
"
$test = Get-SwisData $Swis -Query $SWISConnectionTest


$Swis = NWG-Orion-ConnectToSwis-Trusted -OrionServer $OrionServer -SecondaryOrionServer $SecondaryOrionServer

## Connect to SQL
CD SQLServer:\SQL\$SQLServer\PMGBVSQL01
CD databases\$DBName\tables\

Write-To-Log -Logfile $Log -Level INFO -Message "Starting HW Sensor Threshold Update"
If ($Swis)
{
    Write-To-Log -Logfile $Log -Level INFO -Message "Connected to SWIS"
}


#List of HW Sensors
$Sensors = Get-SwisData $Swis "SELECT Id, NodeID, Name, Value FROM Orion.HardwareHealth.HardwareItem WHERE HardwareCategoryID = '3' AND (Name LIKE 'Gi%' OR Name LIKE 'te%')"

#List of interfaces
$Int = Get-SwisData $Swis "SELECT InterfaceID, NodeID, Name, Status FROM Orion.NPM.Interfaces WHERE NodeID IN (SELECT NodeID FROM Orion.HardwareHealth.HardwareItem) AND Type = '6'"


#region General HWH Sensors

# Bulk disable power sensors that are not required

$SQLQuery = "
SELECT ID, UniqueName, DisplayName, Status, IsDisabled
FROM HWH_HardwareItem
WHERE IsDisabled = '0' AND (UniqueName = 'Power Supply' OR DisplayName LIKE '%power%')
AND (DisplayName LIKE '%Voltage%' OR DisplayName LIKE '%subslot%' OR DisplayName LIKE '%Trans%' OR DisplayName LIKE '%Receive%')
"

$SQLUpdate = "
UPDATE HWH_HardwareItem SET IsDisabled = '1'
WHERE IsDisabled = '0' AND (UniqueName = 'Power Supply' OR DisplayName LIKE '%power%') 
AND (DisplayName LIKE '%Voltage%' OR DisplayName LIKE '%subslot%' OR DisplayName LIKE '%Trans%' OR DisplayName LIKE '%Receive%')
"

$QueryResult = Invoke-Sqlcmd -Query $SQLQuery -QueryTimeout 180 -OutputSqlErrors $true

If ($QueryResult)
{
    Write-To-Log -Logfile $Log -Level INFO -Message "Found $($QueryResult.ID.Count) Sensors that are not required."

    $UpdateResult = Invoke-Sqlcmd -Query $SQLUpdate -QueryTimeout 180 -OutputSqlErrors $true
}

Write-To-Log -Logfile $Log -Level INFO -Message "Exiting"

#endregion


#region Arista HWH

# Bulk disable Arista sensors that are not required

$SQLAristaQuery = "
SELECT ID, UniqueName, DisplayName, Status, IsDisabled
FROM HWH_HardwareItem
WHERE IsDisabled = '0' AND (UniqueName = 'Power Supply' OR DisplayName LIKE '%power%')
AND (DisplayName LIKE '%Voltage%' OR DisplayName LIKE '%subslot%' OR DisplayName LIKE '%Trans%' OR DisplayName LIKE '%Receive%')
"

$SQLAristaUpdate = "
UPDATE HWH_HardwareItem SET IsDisabled = '1'
WHERE IsDisabled = '0' AND (UniqueName = 'Power Supply' OR DisplayName LIKE '%power%') 
AND (DisplayName LIKE '%Voltage%' OR DisplayName LIKE '%subslot%' OR DisplayName LIKE '%Trans%' OR DisplayName LIKE '%Receive%')
"

$AristaQueryResult = Invoke-Sqlcmd -Query $SQLAristaQuery -QueryTimeout 180 -OutputSqlErrors $true

If ($AristaQueryResult)
{
    Write-To-Log -Logfile $Log -Level INFO -Message "Found $($AristaQueryResult.ID.Count) Sensors that are not required."

    $AristaUpdateResult = Invoke-Sqlcmd -Query $SQLAristaUpdate -QueryTimeout 180 -OutputSqlErrors $true
}

Write-To-Log -Logfile $Log -Level INFO -Message "Exiting"
#endregion



## Stop Debug Transcript (If Enabled)
Stop-Transcript