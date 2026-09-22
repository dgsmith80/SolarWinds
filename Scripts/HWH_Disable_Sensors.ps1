$sensors = Get-SwisData $swis @"
SELECT i.HardwareCategoryStatusID, i.HardwareInfo.ID AS [HardwareInfoID], i.UniqueName, i.FullyQualifiedName
FROM Orion.HardwareHealth.HardwareItem i
WHERE i.IsDisabled ='0' AND i.DisplayName LIKE '%DOM%'
"@

ForEach($s in $sensors)
{
    $HardwareInfoID = $s.HardwareInfoID
    $HardwareCategoryStatusID = $s.HardwareCategoryStatusID
    $UniqueName = $s.UniqueName
    $FullName = $s.FullyQualifiedName
        
function HardwareHealthItemKeyXml([int]$HardwareInfoID, [int]$HardwareCategoryStatusID, [string]$UniqueName) { 

    return @" 

    <HardwareHealthItemKey xmlns='http://schemas.solarwinds.com/2007/08/HardwareHealth'> 

        <HardwareInfoID>$HardwareInfoID</HardwareInfoID> 

        <HardwareCategoryStatusID>$HardwareCategoryStatusID</HardwareCategoryStatusID> 

        <UniqueName>$UniqueName</UniqueName> 

    </HardwareHealthItemKey> 

"@ 

} 
    
    $itemKeyXml = HardwareHealthItemKeyXml $sensor.HardwareInfoID $sensor.HardwareCategoryStatusID $sensor.UniqueName 
    $itemKeys = ([xml]"<array>$itemKeyXml</array>").DocumentElement
    Invoke-SwisVerb $swis Orion.HardwareHealth.HardwareItem DisableSensors @( $itemKeys )
    
    $datetime = Get-Date
    $timeutc = $datetime.ToUniversalTime()
    $message = "$timeutc - INFO - Hardware Sensor: $( [string]$FullName ) is not required. Disabled"
    #$message = "$timeutc - INFO - Hardware Sensor ID: $( $InterfaceID ) : $( $FullName ) : Set to CriticalInterface = $( $Critical )"
    $message >> $Log  #Saves output in log file

}

## Log Entry
Write-To-Log -Logfile $Log -Level INFO -Message "Finished HW Sensor Threshold Update"