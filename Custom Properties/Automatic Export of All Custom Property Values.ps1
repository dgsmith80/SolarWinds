<##################################################################################################################################

Script:                 Export Custom Properties

Author:                 David Smith (THWACK: dgsmith80)

Date created:           27.09.2022
Date Edited:            14.06.2024

Purpose:                Export SolarWinds Platform Custom Properties

Prerequisites:          SwisPowerShell (SolarWinds Orion Module) https://www.powershellgallery.com/packages/SwisPowerShell
                        NWG_SolarWinds Registry Key

##################################################################################################################################>

## Start Debug Transcript  - This is used to troubleshoot scripting issues
#Start-Transcript -Path "D:\Prosperon\SolarWinds Automation\Custom Properties\cp_export_transcript.log" -Append -Force -NoClobber ## Remove leading '#' to enable debug

## Load Modules
#Import-Module SQLServer  ## SQL Server Module not required for this script.
Import-Module SwisPowerShell  ## SwisPowerShell Module (Used to execute SWQL Commands)
Import-Module "D:\SolarWinds\Automation\PSModules\NWG-Orion.psm1" -Force  ## NWG Orion Module (Common Function Library)

## File Locations
$Log = "D:\SolarWinds\Automation\Logging\cp_export_log.txt" ## This is the script log file
$Path = "D:\SolarWinds\Automation\Production\Custom Properties" ## This is where the script is located
$RegistryKey = "HKLM:\SOFTWARE\WOW6432Node\NWG_Solarwinds\" ## The registry contains common DNS Entries

## SWIS Connection Info
$OrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MLV
$SecondaryOrionServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-MAC
$Swis = NWG-Orion-ConnectToSwis-Trusted -OrionServer $OrionServer -SecondaryOrionServer $SecondaryOrionServer

## SQL Connection
#$SQLServer = Get-ItemPropertyValue -Path $RegistryKey -Name NWG-SolarWinds-SQL-Orion
#$DBName = 'SolarWindsOrion'
#CD SQLServer:\SQL\$SQLServer\DMGBSQL01
#CD databases\$DBName\tables\BSQL01
#CD databases\$DBName\tables\


### Test SWIS Connection

 $SWISConnectionTest = "
    SELECT Name, InstanceType, InstanceSiteID 
    FROM Orion.SwisFeature
"
$test = Get-SwisData $Swis -Query $SWISConnectionTest


### File Export

$exportDir = "\\rbsres01\shareddata\Teleco\Data Network Operations\Projects\SolarWinds\Extracts\Custom Property Export"
$exportFileNamePre = "SolarWinds-cp-export"
$exportFileNameStart = $exportDir + "\" + $exportFileNamePre
[string] $fileDateStamp = Get-Date -Format "yyyy-MM-dd"


### Get Active Primary Engine ID

$eQuery = @"
SELECT e.EngineID
FROM Orion.Engines AS e
WHERE e.ServerType = 'Primary'
"@

$mpeID = Get-SwisData -SwisConnection $swis -Query $eQuery


### Log Entry Update

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Starting nightly entity custom properties export."
#write-host "$message"
$message >> $Log  #Saves output in log file

New-SwisObject -SwisConnection $swis -EntityType Orion.Events -Properties @{ EventTime = $timeutc; EngineID = $mpeID; Message = "Starting nightly entity custom properties export."; EventType = 1000; Acknowledged = $false }


### Get Custom Property Data

$customPropertySourceQuery = @"
    SELECT
         cp.Field
        ,CASE WHEN cp.TargetEntity = 'Orion.AlertConfigurationsCustomProperties' THEN 'Orion.AlertConfigurations'
              WHEN cp.TargetEntity = 'Orion.APM.ApplicationCustomProperties' THEN 'Orion.APM.Application'
              WHEN cp.TargetEntity = 'Orion.GroupCustomProperties' THEN 'Orion.Groups'
              WHEN cp.TargetEntity = 'Orion.NodesCustomProperties' THEN 'Orion.Nodes'
              WHEN cp.TargetEntity = 'Orion.NPM.InterfacesCustomProperties' THEN 'Orion.NPM.Interfaces'
              WHEN cp.TargetEntity = 'Orion.ReportsCustomProperties' THEN 'Orion.Report'
              WHEN cp.TargetEntity = 'Orion.SEUM.TransactionCustomProperties' THEN 'Orion.SEUM.Transactions'
              WHEN cp.TargetEntity = 'Orion.VolumesCustomProperties' THEN 'Orion.Volumes'
              ELSE 'Error'
         END AS [TargetEntityName]
    FROM Orion.CustomProperty AS cp
"@

$customPropertySourceData = Get-SwisData -SwisConnection $swis -Query $customPropertySourceQuery


### Create Data Segments

$uniqueEntities = $customPropertySourceData | Where-Object { $_.TargetEntityName -ne "Error" } | Select-Object -Property TargetEntityName | Sort-Object -Unique -Property TargetEntityName

ForEach($entity in $uniqueEntities) {

    $entityCustomProperties = $customPropertySourceData | Where-Object { $_.TargetEntityName -eq $entity.TargetEntityName }

    $entityName = $entity.TargetEntityName

    $totalCPs = $entityCustomProperties.Length

    $countCPs = 0

    $customPropertyValuesQuery = $null

    $customPropertyValuesQuery = "SELECT`n`t entity.Uri"

    $customPropertyValuesQueryEnd = "`nFROM " + $entityName + " AS entity"

    ForEach($cp in $entityCustomProperties) {

        $cpName = $cp.Field

        $customPropertyValuesQuery += "`n`t,entity.CustomProperties.$cpName"

        $countCPs ++

    }

    $customPropertyValuesQuery += $customPropertyValuesQueryEnd

    Write-Output("Query for " + $entityName + " is:`n`n" + $customPropertyValuesQuery + "`n`n")

    $customPropertyDataOutput = Get-SwisData -SwisConnection $swis -Query $customPropertyValuesQuery

    $entityFileName = $entityName -replace "\.","-"

    $exportFile = $exportFileNameStart + "-" + $entityFileName + "-" + $fileDateStamp + ".csv"

    $propertySelect = $null

    $propertySelect = @()

    $propertySelect += "Uri"

    $propertySelectCount = 0

    ForEach($field in $entityCustomProperties.Field) {

        $propertySelect += [string] $field

        $propertySelectCount ++

    }

    $customPropertyDataOutput | Select-Object -Property $propertySelect | Export-Csv -Path $exportFile

}

$matchExportCsv = $exportFileNamePre + '.*\.csv$'

Get-ChildItem -Path $exportDir | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-60) -and $_.Name -match $matchExportCsv } | Remove-Item -Force

$utcDate = (Get-Date).ToUniversalTime()


### Log Entry Update

$datetime = Get-Date             
$timeutc = $datetime.ToUniversalTime()   
$message = "$timeutc - INFO - Completed nightly entity custom properties export."
#write-host "$message"
$message >> $Log  #Saves output in log file

New-SwisObject -SwisConnection $swis -EntityType Orion.Events -Properties @{ EventTime = $utcDate; EngineID = $mpeID; Message = "Completed nightly entity custom properties export."; EventType = 1000; Acknowledged = $false }


### Stop Debug Transcript (If Enabled)

#Stop-Transcript
