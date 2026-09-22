# Script to change the Windows credential set used by existing nodes
# to an alternate set of credentials

#Setup the connection to the Solarwinds server
$npmserver = '127.0.0.1'
$swis = Connect-Swis -Hostname $npmserver -Trusted

#Update these fields to match the Credential Names for the old and new credential sets
$oldcred = '<OldCredentialName>'
$newcred = '<NewCredentialName>'

#These variables limit the credentials to only Windows credentials used by NPM
$credtype = 'SolarWinds.Orion.Core.SharedCredentials.Credentials.UsernamePasswordCredential'
$credowner = 'Orion'

#Query the Credentials table to obtain the CredentialID for the credential sets identified above
$oldcredid = Get-SwisData -SwisConnection $swis -query {"SELECT ID FROM Orion.Credential WHERE Name = @cred AND CredentialType = @ct AND CredentialOwner = @owner"} -parameters @{ cred = $oldcred; ct = $credtype; owner = $credowner }
$newcredid = Get-SwisData -SwisConnection $swis -query {"SELECT ID FROM Orion.Credential WHERE Name = @cred AND CredentialType = @ct AND CredentialOwner = @owner"} -parameters @{ cred = $newcred; ct = $credtype; owner = $credowner }

#Get the URI of the NodeSettings entry for each node using the old credentials
$NodeSettings = Get-SwisData $swis {"SELECT Uri from Orion.NodeSettings WHERE SettingName = 'WMICredential' AND SettingValue = @credid"} -Parameters @{ credid = $oldcredid }

#For each URI collected, replace the credential in the SettingValue column with the new CredentialID
#which will associate the relevant node with the new credential set
Foreach ($ns in $NodeSettings) { Set-SwisObject $swis -Uri $ns -Properties @{ SettingValue = $newcredid }
