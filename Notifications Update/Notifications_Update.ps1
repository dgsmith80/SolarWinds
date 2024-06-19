<##################################################################################################################################

Script:              Notifications Update

Author:              David Smith (THWACK: dgsmith80)
Date created:        27.09.2022
Purpose:             Bulk update all Notifications Custom HTML Widgets

Prerequisites:       SwisPowerShell (SolarWinds Orion Module) https://www.powershellgallery.com/packages/SwisPowerShell

##################################################################################################################################>

### Start Debug Transcript (Remove # to enable)

#Start-Transcript -Path "D:\SolarWinds\Automation\Logging\notification_update_transcript.log" -Append -Force -NoClobber

### File/Registry Locations

$Log = 'D:\SolarWinds\Automation\Logging\NotificationsUpdate.log'
$Path = "D:\SolarWinds\Automation\Production\Notifications Update"
$RegistryKey = 'HKLM:\SOFTWARE\WOW6432Node\NWG_Solarwinds\'

### Load Modules

Import-Module SQLServer
#Write-Host("SQLServer Module Imported")
Import-Module SwisPowerShell
#Write-Host("SwisPowerShell Module Imported")
Import-Module "D:\SolarWinds\Automation\PSModules\NWG-Orion.psm1" -Force
#Write-Host("NWG Orion Module Imported")

### Server Details

## SWIS Connection Info
$OrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MLV
$SecondaryOrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MAC
$Swis = NWG-Orion-ConnectToSwis-Trusted -OrionServer $OrionServer -SecondaryOrionServer $SecondaryOrionServer

## SQL Connection
#$SQLServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-SQL-Orion
#$DBName = 'SolarWindsOrion'
#CD SQLServer:\SQL\$SQLServer\DMGBSQL01
#CD databases\$DBName\tables\

### SWIS Connection Test

 $SWISConnectionTest = "
    SELECT Name, InstanceType, InstanceSiteID 
    FROM Orion.SwisFeature
"
$test = Get-SwisData $Swis -Query $SWISConnectionTest

### Import Text File containing Custom HTML Update

$NewHTML = Get-Content -Path $Path\NotificationsHTML.txt 

### List Resources that Contain Notification Custom HTML Resource

$resourcequery =@"
SELECT
     r.ResourceName
    ,r.ResourceTitle
    ,r.ResourceID
    ,r.ViewID
    ,r.Uri AS [ResourceUri]
    ,rp.PropertyName
    ,rp.PropertyValue
    ,rp.Uri AS [ResourcePropertyUri]
FROM Orion.Resources AS r
LEFT JOIN Orion.ResourceProperties AS rp ON rp.ResourceID = r.ResourceID AND rp.PropertyName = 'HTML'
WHERE r.ResourceName = 'Custom HTML' AND r.ResourceTitle = 'Notifications'
"@

$resources = Get-SwisData $Swis -Query $resourcequery

### Log Entry Update

$NumOfResources = $resources.count
if ($NumOfResources -eq $null)
{
    $NumOfResources = 0
}

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Number of Notification Resources found that require updating: " + $NumOfResources
#write-host "$message"
#Saves output in log file
$message >> $Log

### Update the HTML Content of the Widget

$UpdatedResourceCount = 0

    foreach($r in $resources)
    {
    $id = $r.ResourceID
    $view = $r.ViewID
    $RPUri = $r.ResourcePropertyUri
    
    #Set-SwisObject $swis -Uri $RPUri -Properties @{PropertyValue = $NewHTML}
    
    $datetime = Get-Date             
    $timeutc = $datetime.ToUniversalTime()   
    $message = "$timeutc - INFO - Updated ResourceID: $id on ViewID: $view"
    write-host "$message"
    #saves output in log file
    $message >> $Log

    $UpdatedResourceCount++
    }

### Log Entry Update

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Script updated $UpdatedResourceCount Notification Resources."
write-host "$message"
#saves output in log file
$message >> $Path

### Stop Debug Transcript (If Enabled)

#Stop-Transcript