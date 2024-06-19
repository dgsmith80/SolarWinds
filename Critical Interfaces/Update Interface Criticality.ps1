<##################################################################################################################################

Script:					Update Interface Criticality

Author:              	David Smith (THWACK: dgsmith80)
Date created:        	08.11.2023
Date Last Updated:		14.06.2024
Purpose:             	Update CriticalInterface Custom Property based on various conditions.

Prerequisites:       	SwisPowerShell (SolarWinds Orion Module) https://www.powershellgallery.com/packages/SwisPowerShell
                       	NWG_SolarWinds Registry Key

##################################################################################################################################>

### Start Debug Transcript (Remove ## to enable)
#Start-Transcript -Path "D:\SolarWinds\Automation\Logging\critical_interfaces_transcript.log" -Append -Force -NoClobber


### File/Registry Locations
$Log = "D:\SolarWinds\Automation\Logging\critical_interfaces_log.txt" ## This is the script log file
$Path = "D:\SolarWinds\Automation\Production\Critical Interfaces" ## This is where the script is located
$RegistryKey = "HKLM:\SOFTWARE\WOW6432Node\NWG_Solarwinds\" ## The registry contains common DNS Entries


### Load Modules
#Import-Module SQLServer  ## SQL Server Module not required for this script.
Import-Module SwisPowerShell  ## SwisPowerShell Module (Used to execute SWQL Commands)
Import-Module "D:\SolarWinds\Automation\PSModules\NWG-Orion.psm1" -Force  ## NWG Orion Module (Common Function Library)


### Server Details
## SWIS Connection Info
$OrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MLV # This attempts connection to the MLV 
$SecondaryOrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MAC
$Swis = NWG-Orion-ConnectToSwis-Trusted -OrionServer $OrionServer -SecondaryOrionServer $SecondaryOrionServer

## SQL Connection (Not required for this script)
#$SQLServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-SQL-Orion
#$DBName = 'SolarWindsOrion'
#CD SQLServer:\SQL\$SQLServer\DMGBSQL01
#CD databases\$DBName\tables\BSQL01
#CD databases\$DBName\tables\


### Test SWIS Connection
 $SWISConnectionTest = "SELECT Name, InstanceType, InstanceSiteID FROM Orion.SwisFeature"

$test = Get-SwisData $Swis -Query $SWISConnectionTest


### Get Active Primary Engine ID
$eQuery = "SELECT e.EngineID FROM Orion.Engines AS e WHERE e.ServerType = 'Primary'"

$mpeID = Get-SwisData -SwisConnection $swis -Query $eQuery

### Log Entry Update

$datetime = Get-Date -Format "yyyy-MM-dd HH:mm"
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Starting Critical Interface Update Process."
$message >> $Log  #Saves output in log file

New-SwisObject -SwisConnection $swis -EntityType Orion.Events -Properties @{ EventTime = $timeutc; EngineID = $mpeID; Message = "Starting Interface Criticality Update Process."; EventType = 1000; Acknowledged = $false }

#region MONC & MONL Tags

### Get Critical Interfaces containing MONC & MONL Tags

$monc_monl_query = @"
SELECT icp.InterfaceID, icp.Interface.FullName, icp.CriticalInterface, icp.Uri
FROM Orion.NPM.InterfacesCustomProperties icp
WHERE icp.CriticalInterface = '0' AND icp.Interface.Status = '1'
AND (icp.Interface.InterfaceAlias LIKE '%MONC%' OR icp.Interface.InterfaceAlias LIKE '%MONL%')
"@

$monc_monl_interfaces = Get-SwisData -SwisConnection $swis -Query $monc_monl_query

## Set Critical Value to True
$crittrue = '1'


### Interface Count & Log

$numOfInterfaces = $monc_monl_interfaces.Count
if ($numOfInterfaces -eq $null)
{
    $numOfInterfaces = 0
}

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Number of MONC & MONL Interfaces found with CriticalInterface as False: " + $numOfInterfaces
$message >> $Log  #Saves output in log file

$updatedInterfaceCount = 0


### Each Loop 

if($numOfInterfaces -gt 0)
{
    foreach($i in $monc_monl_interfaces)
    {
        $interfaceID = $i.InterfaceID
        $FullName = $i.FullName
        $Critical = $i.CriticalInterface
        $IntUri = $i.Uri
        
        write-host "$interfaceID -- $Fullname -- True"
        Set-SwisObject $swis -Uri $IntUri -Properties @{CriticalInterface = $crittrue}

        $datetime = Get-Date             
        $timeutc = $datetime.ToUniversalTime()
        $message = "$timeutc - INFO - Interface: ID: $( $InterfaceID ) : $( $FullName ) : Set to CriticalInterface = True."
        $message >> $Log  #Saves output in log file

        $updatedInterfaceCount++
    }    
}


### Log Entry Update

$datetime = Get-Date
$timeutc = $datetime.ToUniversalTime()
$message = "$timeutc - INFO - Completed Critical Interface Update Process for $updatedInterfaceCount MONC & MONL Interfaces."
$message >> $Log  #Saves output in log file

#endregion

#region NTMS Devics

### Get Critical Interfaces for NTMS Devices

$ntms_query = @"
SELECT icp.InterfaceID, icp.Interface.FullName, icp.CriticalInterface, icp.Uri
FROM Orion.NPM.InterfacesCustomProperties icp
WHERE icp.CriticalInterface = '0' AND icp.Interface.Status = '1'
AND icp.Interface.Node.CustomProperties.Team = 'NTMS'
AND icp.Interface.Node.CustomProperties.Priority IN ('P1', 'P2')
"@

$ntms_interfaces = Get-SwisData -SwisConnection $swis -Query $ntms_query

## Set Critical Value to True
$crittrue = '1'


### Interface Count & Log

$numOfInterfaces = $ntms_interfaces.Count
if ($numOfInterfaces -eq $null)
{
    $numOfInterfaces = 0
}

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Number of NTMS Interfaces found with CriticalInterface as False: " + $numOfInterfaces
$message >> $Log  #Saves output in log file

$updatedInterfaceCount = 0

### Each Loop 

if($numOfInterfaces -gt 0)
{
    foreach($i in $ntms_interfaces)
    {
        $interfaceID = $i.InterfaceID
        $FullName = $i.FullName
        $Critical = $i.CriticalInterface
        $IntUri = $i.Uri
        
        Set-SwisObject $swis -Uri $IntUri -Properties @{CriticalInterface = $crittrue}
        write-host "$interfaceID -- $Fullname -- True"

        $datetime = Get-Date             
        $timeutc = $datetime.ToUniversalTime()
        $message = "$timeutc - INFO - Interface: ID: $( $InterfaceID ) : $( $FullName ) : Set to CriticalInterface = True."
        $message >> $Log  #Saves output in log file

        $updatedInterfaceCount++
    }    
}


### Log Entry Update

$datetime = Get-Date
$timeutc = $datetime.ToUniversalTime()
$message = "$timeutc - INFO - Completed Critical Interface Update Process for $updatedInterfaceCount NTMS Interfaces."
$message >> $Log  #Saves output in log file

#endregion

#region Security Interfaces

### Get Critical Interfaces for Security Devices

$security_query = @"
SELECT icp.InterfaceID, icp.Interface.FullName, icp.CriticalInterface, icp.Uri
FROM Orion.NPM.InterfacesCustomProperties icp
WHERE icp.CriticalInterface = '0' AND icp.Interface.Status = '1'
AND icp.Interface.Node.CustomProperties.Team = 'Security'
AND icp.Interface.Caption NOT LIKE '%VPN%'
AND icp.Interface.Caption NOT LIKE '%SSL%'
AND icp.Interface.Caption NOT LIKE '%VLAN%'
"@

$security_interfaces = Get-SwisData -SwisConnection $swis -Query $security_query

## Set Critical Value to True
$crittrue = '1'


### Interface Count & Log

$numOfInterfaces = $security_interfaces.Count
if ($numOfInterfaces -eq $null)
{
    $numOfInterfaces = 0
}

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Number of Security Interfaces found with CriticalInterface as False: " + $numOfInterfaces
$message >> $Log  #Saves output in log file

$updatedInterfaceCount = 0

### Each Loop 

if($numOfInterfaces -gt 0)
{
    foreach($i in $security_interfaces)
    {
        $interfaceID = $i.InterfaceID
        $FullName = $i.FullName
        $Critical = $i.CriticalInterface
        $IntUri = $i.Uri
        
        write-host "$interfaceID -- $Fullname -- True"
        Set-SwisObject $swis -Uri $IntUri -Properties @{CriticalInterface = $crittrue}

        $datetime = Get-Date             
        $timeutc = $datetime.ToUniversalTime()
        $message = "$timeutc - INFO - Interface: ID: $( $InterfaceID ) : $( $FullName ) : Set to CriticalInterface = True."
        $message >> $Log  #Saves output in log file

        $updatedInterfaceCount++
    }    
}


### Log Entry Update

$datetime = Get-Date
$timeutc = $datetime.ToUniversalTime()
$message = "$timeutc - INFO - Completed Critical Interface Update Process for $updatedInterfaceCount Security Interfaces."
$message >> $Log  #Saves output in log file

#endregion


#region MON-IO Removal

### Get MON-IO Interfaces for All Interfaces flagged as Critical

$monio_query = @"
SELECT icp.InterfaceID, icp.Interface.FullName, icp.CriticalInterface, icp.Uri
FROM Orion.NPM.InterfacesCustomProperties icp
WHERE icp.CriticalInterface = '1' AND icp.Interface.Status = '1'
AND icp.Interface.InterfaceAlias LIKE '%MON-IO%'
"@

$monio_interfaces = Get-SwisData -SwisConnection $swis -Query $monio_query

## Set Critical Value to False
$critfalse = '0'


### Interface Count & Log

$numOfInterfaces = $monio_interfaces.Count
if ($numOfInterfaces -eq $null)
{
    $numOfInterfaces = 0
}

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Number of MON-IO Interfaces found with CriticalInterface as True: " + $numOfInterfaces
$message >> $Log  #Saves output in log file

$updatedInterfaceCount = 0

### Each Loop 

if($numOfInterfaces -gt 0)
{
    foreach($i in $monio_interfaces)
    {
        $interfaceID = $i.InterfaceID
        $FullName = $i.FullName
        $Critical = $i.CriticalInterface
        $IntUri = $i.Uri
        
        write-host "$interfaceID -- $Fullname -- False"
        Set-SwisObject $swis -Uri $IntUri -Properties @{CriticalInterface = $critfalse}

        $datetime = Get-Date             
        $timeutc = $datetime.ToUniversalTime()
        $message = "$timeutc - INFO - Interface: ID: $( $InterfaceID ) : $( $FullName ) : Set to CriticalInterface = False."
        $message >> $Log  #Saves output in log file

        $updatedInterfaceCount++
    }    
}


### Log Entry Update

$datetime = Get-Date
$timeutc = $datetime.ToUniversalTime()
$message = "$timeutc - INFO - Completed Critical Interface Update Process for $updatedInterfaceCount MON-IO Interfaces."
$message >> $Log  #Saves output in log file

#endregion


### Event Log Completion Update

New-SwisObject -SwisConnection $swis -EntityType Orion.Events -Properties @{ EventTime = $utcDate; EngineID = $mpeID; Message = "Completed Interface Criticality Update Process."; EventType = 1000; Acknowledged = $false }


## Stop Debug Transcript (If Enabled)

#Stop-Transcript