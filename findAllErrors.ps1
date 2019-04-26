###############################################################################
# Find All COM errors
# Author: Kit Menke
# Version 1.0 11/6/2016
# Modified: John Anson 2019-04-26
###############################################################################

param (
	[alias('days')] [int]$maxDays=9999 # number of days to search back in the event log
)

# Notes:
# Get-EventLog doesn't quite work I guess:
# https://stackoverflow.com/questions/31396903/get-eventlog-valid-message-missing-for-some-event-log-sources#
# Get-EventLog Application -EntryType Error -Source "DistributedCOM"
# The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID
#$logs = Get-EventLog -LogName "System" -EntryType Error -Source "DCOM" -Newest 1 -Message "The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID*"

$EVT_MSG = "The application-specific permission settings do not grant Local Activation permission for the COM Server application with CLSID"
# Search for System event log ERROR entries starting with the specified EVT_MSG
# Level 2 is error, 3 is warning
$logEntries = @(Get-WinEvent -FilterHashTable @{LogName='System'; Level=2; Id=10016} | Where-Object { $_.Message -like "$EVT_MSG*" } )
if (!$logEntries) {
  Write-Host "No event log entries found."
  exit 1
}

write-host "$($logEntries.count) entries found"

$ClsApps = New-Object System.Collections.ArrayList
$count = 0
$dupCount = 0
$now = get-date
foreach ($logEntry in $logEntries) {

  if ($logEntry.timeCreated -lt $now.adddays(-$maxDays)) {
    write-host "# entry time ($($logEntry.timeCreated))is more than $maxDays days ago. Quit"
    break
  }
  
  # Write-Host "Found an event log entry :"
  # Write-Host ($logEntry | Format-List | Out-String)
  # Write-Host ($logEntry.Properties | Format-List | Out-String)

  # Get CLSID and APPID from the event log entry
  # which we'll use to look up keys in the registry
  $CLSID = $logEntry.Properties[3].Value
  # Write-Host "CLSID=$CLSID"
  $APPID = $logEntry.Properties[4].Value
  # Write-Host "APPID=$APPID"
  $USERDOMAIN = $logEntry.Properties[5].Value
  # Write-Host "USERDOMAIN=$USERDOMAIN"
  $USERNAME = $logEntry.Properties[6].Value
  # Write-Host "USERNAME=$USERNAME"
  $USERSID = $logEntry.Properties[7].Value
  # Write-Host "USERSID=$USERSID"

  # Skip duplicates
  $ClsApp = "$CLSID $APPID"
  if ($ClsApps.contains($ClsApp)) {
    # write-host "# $ClsApp already found"
    $dupCount++
    continue
  }
  $count++
  $ClsApps.add($ClsApp) > $null
  write-host "# $($logEntry.timeCreated)"
  Write-Host ".\checkerrors.ps1 ""$APPID"" ""$USERSID"""
  Write-Host ".\fixerrors.ps1 ""$APPID"" ""$CLSID"" ""$USERDOMAIN"" ""$USERNAME"""
}

write-host "# $count unique entries found"
write-host "# $dupCount duplicates skipped"
