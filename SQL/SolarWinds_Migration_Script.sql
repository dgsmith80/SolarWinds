DECLARE @oldHostname nvarchar(max)
SET @oldHostname = 'Server1'
DECLARE @newHostname nvarchar(max)
SET @newHostname = 'Server2'

UPDATE Engines SET ServerName = @newHostname WHERE ServerName = @oldHostname
UPDATE WebSettings SET SettingValue = @newHostname WHERE SettingName = 'JobSchedulerHost'
UPDATE Websites SET ServerName = @newHostname WHERE ServerName = @oldHostname
UPDATE OrionServers SET HostName = @newHostname WHERE HostName = @oldHostname
