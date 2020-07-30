<#
.NOTES
	Name: Generate-MessageProfile.ps1
	Author: Daniel Sheehan
	Requires: PowerShell v2 or higher and to be run through the full Exchange
	Management Shell (not a remote Shell session). The account running this
	script needs to have administrator rights on the Exchange servers and rights
	to query tracking logs and mailboxes.
	Version 2.0 - 12/30/2016: Introduced multi-threading to allow for data
	gathering form multiple servers simultaneously. Reconfigured server retry
	mechanism to add retries at the end of the server job list. Added override
	mechanism to allow a site message profile to be created even when a
	percentage of its servers are inaccessible/don't return data.
	Version 2.1 - 1/25/2017: Fixed an issue with a site not being skipped if it
	had no recorded messages in it.
	*** For a complete version history, visit the script's Link below. ***
	############################################################################
	The sample scripts are not supported under any Microsoft standard support
	program or service. The sample scripts are provided AS IS without warranty
	of any kind. Microsoft further disclaims all implied warranties including,
	without limitation, any implied warranties of merchantability or of fitness
	for a particular purpose. The entire risk arising out of the use or
	performance of the sample scripts and documentation remains with you. In no
	event shall Microsoft, its authors, or anyone else involved in the creation,
	production, or delivery of the scripts be liable for any damages whatsoever
	(including, without limitation, damages for loss of business profits,
	business interruption, loss of business information, or other pecuniary
	loss) arising out of the use of or inability to use the sample scripts or
	documentation, even if Microsoft has been advised of the possibility of such
	damages.
	############################################################################
.SYNOPSIS
	Generates a user message profile (used for Exchange server sizing and other
	efforts), based upon the specified date range and other optional parameters,
	for each specified Exchange site.
.DESCRIPTION
	This script enumerates all of the Exchange Mailbox and Hub Transport servers
	in each specified AD site(s), and then loops through each one gathering the
	count of mailboxes and also messages that have been sent to or received from
	a mailbox during the specified date range. Specific types of data can be
	excluded from the data gathering process by using various script parameters.
	The gathered information is then compiled into a table organized by site,
	and is either stored in memory or optionally exported to CSV file.
	Message profiles stored in multiple CSV files, either from separate site
	collections and/or from the same site collections over a period of time, can
	be imported to provide an aggregated message profile.
.PARAMETER ADSites
	This optional parameter for the Parameter Set "Gather" defaults to "*" which
	indicates all AD sites with Exchange should be processed. Alternatively,
	explicit site names, site names with wild cards, or any combination thereof
	can be used to specify multiple AD sites to filter on. The format for
	multiple sites is each site name in quotes, separated by a comma with no
	spaces such as:
	"Site1","Site2","AltSite*", etc...
.PARAMETER StartOnDate
	This mandatory parameter for the Parameter Set "Gather" specifies the date
	(at 12:00AM) the message tracking log search should start on.
	The format is MM/DD/YYYY.
.PARAMETER EndBeforeDate
	This mandatory parameter for the Parameter Set "Gather" specifies the date
	(at 12:00AM) the message tracking log search should end before. This means
	that if the desired search window is Monday through Friday, Saturday needs
	to be specified so the search "ends before" (stops) at 12:00AM Saturday.
	This will allow for all of Friday to be included in the search.
	The format is MM/DD/YYYY.
.PARAMETER ExcludeHealthData
	This optional switch for the Parameter Set "Gather" excludes messages to or
	from Managed Availability "HealthMailbox" and the older SCOM "extest_"
	mailboxes, which could artificially inflate the message profile for a site.
.PARAMETER ExcludeJournalData
	This optional switch for the Parameter Set "Gather" excludes journal
	messages from the data collection. By default messages delivered to journal
	mailboxes will be included with the message profile, which could
	artificially inflate the message profile for a site.
.PARAMETER ExcludePFData
	This optional switch for the Parameter Set "Gather" attempts to filter out
	messages sent to or from legacy Exchange 2007/2010 Public Folder databases.
	This is not needed if there are no legacy Exchange Public Folder databases.
.PARAMETER ExcludeRoomMailboxes
	This optional switch for the Parameter Set "Gather" excludes messages
	to or from room mailboxes. By default equipment and discovery mailboxes are
	excluded from the count as they negatively skew the average user message
	profile. Room mailboxes are included by default because they can
	send/receive email.
.PARAMETER BypassRPCCheck
	This optional switch for the Parameter Set "Gather" instructs the script to
	bypass the additional RPC connectivity test to remote computers through
	Get-WMIObject. Basic PING tests are always used to initially test
	connectivity to remote computers. Bypassing the RPC check should not be
	necessary as long as the account running the script has the appropriate
	permissions to connect to WMI on the remote computers.
.PARAMETER MaxServerTries
	This optional parameter for the Parameter Set "Gather" specifies the maximum
	number of times to try to gather data from a server when there are issues
	gathering data. The default value of 3 means the script will try to gather
	data from each server up to 3 times before giving up on it and marking it as
	a skipped server.
.PARAMETER MinServersPercent
	This optional parameter for the Parameter Set "Gather", which defaults to
	100%, specifies the minimum percentage of servers in a site that must be
	accessible and also return data to adequately generate a message profile. If
	this percentage is not met, because too many servers are inaccessible or
	they exceed the MaxServerTries during data gathering, the site is skipped
	(recorded as a SkippedSite so the script can be quickly re-run against it)
	and not included in the final message profile collection.
	The format is a number value without the "%".
	*** It is highly recommended to leave this value at 100, because missing
	even one server could result in a potentially skewed message profile. ***
.PARAMETER MaxThreads
	This optional parameter for the Parameter Set "Gather" specifies the maximum
	number of simultaneous server data gathering jobs (threads). Each job
	increases the memory and CPU load on the server running this script.
	Therefore, the number of jobs defaults to 1/4 (rounded up) of logical cores
	if the system running the script is running Exchange services, or 1/2 if it
	is not. Visit the script's Link below for more information.
	*** Monitor CPU and memory impact and adjust as necessary. ***
.PARAMETER Confirm
	This optional parameter for the Parameter Set "Gather" bypasses the warning
	prompts for changes to the MinServersPercent and MaxThreads parameters.
.PARAMETER ExcludeSites
	This optional parameter for both the "Gather" and "Import" Parameter Sets
	specifies which sites should be excluded from data processing. This is
	useful when you want to use a wild card to gather data from multiple sites,
	but you want to exclude specific sites that would normally be included in
	the wild card collection. Likewise, sites that do not house any user
	mailboxes, such as dedicated Hybrid sites, can be excluded.
	For data importing, this is useful when a site needs to be excluded from a
	previous collection. The format for multiple sites is each individual site
	name in quotes, wild cards are not supported, separated by a comma with no
	spaces such as:
	"Site1","Site2", etc...
.PARAMETER InCSVFile
	This mandatory parameter for the Parameter Set "Import" specifies the path
	and file name of the CSV to import previously collected data from.
.PARAMETER InMemory
	This mandatory switch for the Parameter Set "Existing" instructs the script
	to only use existing in memory data. This intended only to be used with the
	AverageAllSites parameter switch.
.PARAMETER AverageAllSites
	This optional switch instructs the script to create an "~All Sites" entry in
	the collection that represents an average message profile of all sites
	collected. If an existing "~All Sites" entry already exists, its data is
	overwritten with the updated data.
.PARAMETER OutCSVFile
	This optional parameter specifies the path and file name of the CSV to
	export the collected data to. If this parameter is omitted, then the
	collected data is saved in the shell variable $MessageProfile.
.EXAMPLE
	[PS] C:\>.\Generate-MessageProfile.ps1 -StartOnDate 12/1/2014
	-EndBeforeDate 12/6/2014 -ExcludeHealthData -OutCSVFile AllSites.CSV
	Exchange servers in all sites are processed starting on Monday 12/1/2014
	through the end of Friday 12/5/2014. The collected data, which excludes
	message data for Exchange 2013+ HealthMailboxes and any extest_ mailboxes,
	is exported to the AllSites.CSV file.
.EXAMPLE
	[PS] C:\>.\Generate-MessageProfile.ps1 -ADSites East* -StartOnDate 12/1/2014
	-EndBeforeDate 12/2/2014 -Verbose -Debug
	Exchange servers in sites that start with "East" are processed starting on
	Monday 12/1/2014 through the end of Monday 12/1/2014 (I.E. It's data
	gathering for just one day). Output the additional Verbose and Debug
	information to the screen while the script is running. The collected data is
	available in the $MessageProfile variable after the script completes.
.EXAMPLE
	[PS] C:\>.\Generate-MessageProfile.ps1 -ADSites "EastDC1","West*"
	-StartOnDate 12/1/2014 -EndBeforeDate 12/31/2014 -OutCSVFile MultiSites.CSV
	-ExcludePFData -ExcludeJournalData
	Exchange servers in the EastDC1 site and any sites that start with "West"
	are processed starting on Monday 12/1/2014 through the end of Tuesday
	12/30/2014. The collected data, which should exclude most Public Folder
	traffic and all Journal messages, is exported to the MultiSites.CSV file.
.EXAMPLE
	[PS] C:\>.\Generate-MessageProfile.ps1 -InCSVFile .\PreviousCollection.CSV
	The data from the PreviousCollection CSV file in the current working
	directory is imported into the in-memory $MessageProfile data collection for
	future use.
.EXAMPLE
	[PS] C:\>.\Generate-MessageProfile.ps1 -InMemory -AverageAllSites
	The previously collected data stored in the $MessageProfile variable
	is processed and an average for all the sites is added to the data
	collection as the site name "~All Sites".
.LINK
	https://gallery.technet.microsoft.com/Generate-Message-Profile-7d0b1ef4
#>
#Requires -Version 2.0

# Use the CmdletBinding function so the script accepts and understands -Verbose and -Debug and sets the default parameter set to
#   "Gather". The Write-Verbose and Write-Debug statements in this script will activate only if their respective switches are used.
[CmdletBinding(DefaultParameterSetName = "Gather")]
# Read in all the command line parameters, grouping most of them into 3 parameter sets with the exception of the ExcludeSites
#   parameter which is included in two parameter sets.
Param (
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	# Default to the value of "*" for all sites if no site name is specified.
	[Array]$ADSites = "*",
	[Parameter(ParameterSetName = "Gather", Mandatory = $True)]
	[DateTime]$StartOnDate,
	[Parameter(ParameterSetName = "Gather", Mandatory = $True)]
	[DateTime]$EndBeforeDate,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Switch]$ExcludeHealthData,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Switch]$ExcludeJournalData,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Switch]$ExcludePFData,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Switch]$ExcludeRoomMailboxes,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Switch]$BypassRPCCheck,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Int]$MaxServerTries = 3,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Int]$MinServersPercent = 100,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Int]$MaxThreads,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)]
	[Bool]$Confirm = $True,
	[Parameter(ParameterSetName = "Gather", Mandatory = $False)][Parameter(ParameterSetName = "Import", Mandatory = $False)]
	[Array]$ExcludeSites,
	[Parameter(ParameterSetName = "Import", Mandatory = $True)]
	[String]$InCSVFile,
	[Parameter(ParameterSetName = "Existing", Mandatory = $True)]
	[Switch]$InMemory,
	[Parameter(Mandatory = $False)]
	[Switch]$AverageAllSites,
	[Parameter(Mandatory = $False)]
	[String]$OutCSVFile
)

# Start tracking the time this script takes to run.
$StopWatch = New-Object System.Diagnostics.Stopwatch
$StopWatch.Start()

#region SetSpecialColors
# If the -Debug parameter was used, then record the existing Debug preference and then set it to Continue so bypass script pauses for Write-Debug.
If ($PSBoundParameters["Debug"]) {
	$HoldDebugPreference = $DebugPreference
	$DebugPreference = "Continue"
	# Also save the default foreground text color and then change it to Dark Red.
	$DebugForeground = $Host.PrivateData.DebugForegroundColor
	$Host.PrivateData.DebugForegroundColor = "Magenta"
}
# If the -Verbose parameter was used, then save the default foreground text color and then change it to Cyan.
If ($PSBoundParameters["Verbose"]) {
	$VerboseForeground = $Host.PrivateData.VerboseForegroundColor
	$Host.PrivateData.VerboseForegroundColor = "Cyan"
}
#endregion SetSpecialColors

# Function to set Debug/Verbose settings back to default, gracefully shut down the stop watch, report the amount of time the script took to run.
Function _Exit-Script {
	Param (
		[Parameter(Mandatory = $False)]
		[Bool]$HardExit = $False
	)

	# If HardExit is specified then note the script will now exit and issue the EXIT command to exit the script at the end.
	If ($HardExit) {
		Write-Host "The script will now exit."
	}

	#region SetDefaultColors
	# If the DebugForeground variable was set at the top of the script, then set the Debug preference and color back to the default.
	If ($DebugForeground) {
		$DebugPreference = $HoldDebugPreference
		$Host.PrivateData.DebugForegroundColor = $DebugForeground
	}
	# If the VerboseForeground variable was set at the top of the script, then change the color back to the default.
	If ($VerboseForeground) {
		$Host.PrivateData.VerboseForegroundColor = $VerboseForeground
	}
	#endregion SetDefaultColors

	$StopWatch.Stop()
	$ElapsedTime = $StopWatch.Elapsed
	$TotalHours = ($ElapsedTime.Days * 24) + $ElapsedTime.Hours
	Write-Host ""
	Write-Host "The script took $TotalHours hour(s), $($ElapsedTime.Minutes) minute(s), and $($ElapsedTime.Seconds) second(s) to run."

	If ($HardExit) {
		EXIT
	}
}

# If the MessageProfile variable was not defined from a previous script run, then create it as MessageProfile to hold all of the collected table.
If (-not($MessageProfile)) {
	$MessageProfile = New-Object System.Data.DataTable "MessageProfile"
	# Create columns in the DataTable by specifying their name and property type.
	$MessageProfile.Columns.Add("SiteName",[String]) | Out-Null
	$MessageProfile.Columns.Add("Mailboxes",[Int]) | Out-Null
	$MessageProfile.Columns.Add("AvgTotalMsgs",[Int]) | Out-Null
	$MessageProfile.Columns.Add("AvgTotalKB",[Int]) | Out-Null
	$MessageProfile.Columns.Add("AvgSentMsgs",[Int]) | Out-Null
	$MessageProfile.Columns.Add("AvgRcvdMsgs",[Int]) | Out-Null
	$MessageProfile.Columns.Add("AvgSentKB",[Int]) | Out-Null
	$MessageProfile.Columns.Add("AvgRcvdKB",[Int]) | Out-Null
	$MessageProfile.Columns.Add("SentMsgs",[Int64]) | Out-Null
	$MessageProfile.Columns.Add("RcvdMsgs",[Int64]) | Out-Null
	$MessageProfile.Columns.Add("SentKB",[Int64]) | Out-Null
	$MessageProfile.Columns.Add("RcvdKB",[Int64]) | Out-Null
	$MessageProfile.Columns.Add("UTCOffset",[Double]) | Out-Null
	$MessageProfile.Columns.Add("TimeSpan",[Double]) | Out-Null
	$MessageProfile.Columns.Add("TotalDays",[Int]) | Out-Null
	# Set the SiteName column as the unique key so the rows can be searched by site name.
	$MessageProfile.PrimaryKey = $MessageProfile.Columns["SiteName"]
}

# Function to add site profile data to the DataTable which will be called later in the script.
Function _Add-SiteData {
	# Read in the mandatory parameters passed in by their position in the pipeline.
	Param (
		[Parameter(Mandatory = $True, Position = 0)]
		$SiteName,
		[Parameter(Mandatory = $True, Position = 1)]
		$MailBoxes,
		[Parameter(Mandatory = $True, Position = 2)]
		$SentCount,
		[Parameter(Mandatory = $True, Position = 3)]
		$SentSize,
		[Parameter(Mandatory = $True, Position = 4)]
		$ReceivedCount,
		[Parameter(Mandatory = $True, Position = 5)]
		$ReceivedSize,
		[Parameter(Mandatory = $True, Position = 6)]
		$UTCOffset,
		[Parameter(Mandatory = $True, Position = 7)]
		$TimeSpan,
		[Parameter(Mandatory = $True, Position = 8)]
		$TotalDays
	)
	# Check if the site's profile already exists as a row in the DataTable from a previous script run by checking the site name.
	If ($ProfileRow = $MessageProfile.Rows.Find($SiteName)) {
		# The site name was found, so capture the data passed into the function and make the necessary calculations.
		Write-Verbose " * Adding data to the existing $($ProfileRow.TotalDays) days for the `"$SiteName`" in the DataTable."
		$NewTotalDays = $TotalDays
		$NewSentCount = $SentCount
		# Calculate the combined total size of sent messages in KB, rounding to the nearest whole number.
		$NewSentKB = [Math]::Round(($SentSize / 1KB),0,"AwayFromZero")
		$NewReceivedCount = $ReceivedCount
		# Calculate the combined total size of received messages in KB, rounding to the nearest whole number.
		$NewReceivedKB = [Math]::Round(($ReceivedSize / 1KB),0,"AwayFromZero")
		# Only add the existing site data to the data passed into the function if the site name is not "~All Sites" as that is handled separately.
		If ($SiteName -notlike "~All Sites") {
			$NewTotalDays += $ProfileRow.TotalDays
			$NewSentCount += $ProfileRow.SentMsgs
			$NewSentKB += $ProfileRow.SentKB
			$NewReceivedCount += $ProfileRow.RcvdMsgs
			$NewReceivedKB += $ProfileRow.RcvdKB
		}
		# Update the site information in the row.
		$ProfileRow.Mailboxes = $Mailboxes
		$ProfileRow.SentMsgs = $NewSentCount
		$ProfileRow.SentKB = $NewSentKB
		$ProfileRow.RcvdMsgs = $NewReceivedCount
		$ProfileRow.RcvdKB = $NewReceivedKB
		# If the combined SentCount is not 0, then calculate the average sent message size by dividing the sent messages in KB by the number of
		#   sent messages, rounding to the nearest whole number.
		If ($NewSentCount -ne 0) {
			$ProfileRow.AvgSentKB = [Math]::Round(($NewSentKB / $NewSentCount),0,"AwayFromZero")
		}
		# If the combined ReceivedCount is not 0, then calculate the average received message size by dividing the received messages in KB by the
		#   number of received messages rounding to the nearest whole number.
		If ($NewReceivedCount -ne 0) {
			$ProfileRow.AvgRcvdKB = [Math]::Round(($NewReceivedKB / $NewReceivedCount),0,"AwayFromZero")
		}
		# Calculate the new average total message size by adding both the sent and received message sizes in KB, dividing the number of sent and
		#   received messages, rounding to the nearest whole number.
		$ProfileRow.AvgTotalKB = [Math]::Round((($NewSentKB + $NewReceivedKB) / `
			($NewSentCount + $NewReceivedCount)),0,"AwayFromZero")
		# Add the number of average sent messages by dividing all sent messages by the number of days in the query, then divide by the number of
		#   mailboxes in the site, rounding up to the nearest whole number.
		$ProfileRow.AvgSentMsgs = [Math]::Ceiling(($NewSentCount / $NewTotalDays) / $Mailboxes)
		# Add the number of average received messages by dividing all received messages by the number of days in the query, then divide by the
		#   number of mailboxes in the site, rounding up to the nearest whole number.
		$ProfileRow.AvgRcvdMsgs = [Math]::Ceiling(($NewReceivedCount / $NewTotalDays) / $Mailboxes)
		# Calculate the average total number of messages by adding existing and new sent and received message counts, dividing by the number of
		#   new total days, then dividing by the number of mailboxes in the site, rounding up to the nearest whole number.
		$ProfileRow.AvgTotalMsgs = [Math]::Ceiling((($NewSentCount + $NewReceivedCount) / $NewTotalDays) / $Mailboxes)
		$ProfileRow.UTCOffset = $UTCOffset
		$ProfileRow.TimeSpan = $TimeSpan
		$ProfileRow.TotalDays = $NewTotalDays
		Write-Debug "The data modified in the DataTable for the $SiteName site is:`n$($ProfileRow.ItemArray)"
	} Else {
		# The site name was not found so create a new row in the DataTable and add the relevant data.
		$NewProfileRow = $MessageProfile.NewRow()
		$NewProfileRow.SiteName = $SiteName
		$NewProfileRow.Mailboxes = $Mailboxes
		$NewProfileRow.SentMsgs = $SentCount
		# Calculate the total size of sent messages in KB, rounding to the nearest whole number.
		$NewProfileRow.SentKB = [Math]::Round(($SentSize / 1KB),0,"AwayFromZero")
		# if the SentCount was 0, and if so record the AvgSentKB as 0 to avoid a divide by 0 error.
		If ($SentCount -eq 0) {
			$NewProfileRow.AvgSentKB = 0
		# Otherwise calculate the average sent message size by dividing the sent messages in KB by the number of sent messages, rounding to the
		#   nearest whole number.
		} Else {
			$NewProfileRow.AvgSentKB = [Math]::Round((($SentSize / 1KB) / $SentCount),0,"AwayFromZero")
		}
		# Add the number of messages received in the site to the row.
		$NewProfileRow.RcvdMsgs = $ReceivedCount
		# Calculate the total size of received messages in KB, rounding to the nearest whole number.
		$NewProfileRow.RcvdKB = [Math]::Round(($ReceivedSize / 1KB),0,"AwayFromZero")
		# If the ReceivedCount was 0, then record the AvgRcvdKB as 0 to avoid a divide by 0 error.
		If ($ReceivedCount -eq 0) {
			$NewProfileRow.AvgRcvdKB = 0
		# Otherwise calculate the average received message size by dividing the received messages in KB by the number of received messages
		#   rounding to the nearest whole number.
		} Else {
			$NewProfileRow.AvgRcvdKB = [Math]::Round((($ReceivedSize / 1KB) / $ReceivedCount),0,"AwayFromZero")
		}
		# Calculate the average total message size by adding both the sent and received message sizes in KB, dividing the number of send and
		#   received messages, rounding to the nearest whole number.
		$NewProfileRow.AvgTotalKB = [Math]::Round(((($SentSize + $ReceivedSize) / 1KB) / `
			($SentCount + $ReceivedCount)),0,"AwayFromZero")
		# Add the number of average sent messages by dividing all sent messages by the number of days in the query, then dividing by the number
		#   of mailboxes in the site, rounding up to the nearest whole number.
		$NewProfileRow.AvgSentMsgs = [Math]::Ceiling(($SentCount / $TotalDays) / $Mailboxes)
		# Add the number of average received messages by dividing all received messages by the number of days in the query, then dividing by the
		#   number of mailboxes in the site, rounding up to the nearest whole number.
		$NewProfileRow.AvgRcvdMsgs = [Math]::Ceiling(($ReceivedCount / $TotalDays) / $Mailboxes)
		# Calculate the average total number of messages by adding both the sent and received message counts, dividing by the number of days in
		#   the query, then dividing by the number of mailboxes in the site, rounding up to the nearest whole number.
		$NewProfileRow.AvgTotalMsgs = [Math]::Ceiling((($SentCount + $ReceivedCount) / $TotalDays) / $Mailboxes)
		$NewProfileRow.UTCOffset = $UTCOffset
		$NewProfileRow.TimeSpan = $TimeSpan
		$NewProfileRow.TotalDays = $TotalDays
		# Commit the new row to the MessageProfile DataTable.
		$MessageProfile.Rows.Add($NewProfileRow)
		Write-Debug "The data added to the DataTable for the `"$SiteName`" site is:`n$($NewProfileRow.ItemArray)"
	}
}

# Create the Test-Connectivity function to first verify the remote Computer responds to a ping and then RPC access.
Function _Test-Connectivity {
	# Read in the mandatory Computer name.
	Param(
		[Parameter(Mandatory = $True, Position = 0)]	
		[String]$Computer
	)
	# Try to ping the computer 1 time. NOTE: The Test-Connection ping request has a 1 second hard coded timeout.
	If (Test-Connection -ComputerName $Computer -Quiet -Count 1) {
		# The short ping test was successful so the rest of the If statement is skipped.
	# Otherwise 1 ping request didn't work, so try 4 just in case it was temporary issue. If any 1 of the 4 pings succeed, the test is considered
	#   successful.
	} ElseIf (Test-Connection -ComputerName $Computer -Quiet -Count 4) {
		# The long ping test was successful so the rest of the If statement is skipped.
	} Else {
		# Neither the short or long ping test were successful, so return out of the function with the value of False.
		Return $False
	}
	# Since one of the ping tests was successful, otherwise the function would have already returned out, check to see if the script level
	#   BypassRPCCheck was used.
	If (-not($BypassRPCCheck)) {
		# It wasn't so try the following RPC connection test.
		Try {
			# Connect to the remote computer using WMI which not only tests RPC connectivity but admin connectivity permissions.
			# NOTE: WMI calls can hang indefinitely which is why the PING tests are performed first.
			Get-WmiObject Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop
		# Check and see if an error was caught, and if so see if we can identify it as a known error.
		} Catch {
			If ($_.FullyQualifiedErrorId -Match "UnauthorizedAccessException") {
				# The UnauthorizedAccessException error was found so report it was a permissions issue.
				Write-Host ""
				Write-Host -ForegroundColor Red "You do not have permission to remotely connect to $Computer."
				Write-Warning "The optional script switch -BypassRPCCheck will bypass this part of the connectivity test."
			} Else {
				# Otherwise the error is not known so output it to the screen for a more detailed analysis.
				Write-Host ""
				Write-Host -ForegroundColor Red "There was an error remotely connecting to $Computer with the error code:"
				Write-Host -ForegroundColor Red "$($_.Exception)"
			}
			# Since an error was caught, and the error was already reported to the screen, return out of the function with the value of false.
			Return $False
		}
	}
	# Otherwise the ping test and the WMI test were both successful, so return out of the function with the value of True.
	Return $True
}

# Check to see if the Gather ParameterSet was used.
If ($PsCmdlet.ParameterSetName -like "Gather") {
	# It was so execute the data gathering section of the script.

	# Validate the script is being run in an Exchange Management Shell by looking for the $ExScripts variable, and exit if it is not.
	If (-Not($ExScripts)) {
		Write-Host -ForegroundColor Red "The Exchange Management Shell (EMS) wasn't detected. Please run the script through the EMS."
		_Exit-Script -HardExit $True
	}

	$TodaysDate = Get-Date
	Write-Verbose "Starting Exchange site data gathering at $($TodaysDate.ToString())."

	# If there are any residual jobs from a previous run of this script that terminated abnormally, report them and remove them.
	If ($OldJobs = (Get-Job | Where-Object {$_.Name -like "MessageProfile*"})) {
		Write-Host ""
		Write-Warning ("There is/are $(($OldJobs | Measure-Object).Count) old message profile data gathering job(s) detected on this " + `
			"system that will now be removed.")
		Write-Host "The only times jobs should be left over on this system is if the script previously terminated abnormally, as it is" `
			"designed to clean up after itself."

		ForEach ($OldJob in $OldJobs) {
			Remove-Job $OldJob -Force -Confirm:$False
		}
	}

	#region ValidateInput
	# Validate the StartOnDate and EndBeforeDate variables by comparing them against today's date.
	$StartOnDateOffset = $StartOnDate - $TodaysDate
	$EndBeforeDateOffset = $EndBeforeDate - $TodaysDate
	# If the StartOnDateOffset is greater than or equal to negative one day's worth of "ticks", which means the StartOnDate occurs on or after
	#   today, report that and exit out of the script.
	If ($StartOnDateOffset -ge -864000000000) {
		Write-Debug "The StartOnDate offset is $($StartOnDateOffset.Days) days and $($StartOnDateOffset.Hours) hours from today."
		Write-Host ""
		Write-Host -ForegroundColor Red "The StartOnDate of $($StartOnDate.ToShortDateString()) needs to be changed to a day prior to today's" `
			"date of $($TodaysDate.ToShortDateString())."
		_Exit-Script -HardExit $True
	# Otherwise if the EndBeforeDateOffset shows the End date is past today, report that and exit out of the script.
	} ElseIf ($EndBeforeDateOffset -ge 0) {
		Write-Debug "The EndBeforeDate offset is $($EndBeforeDateOffset.Days) days and $($EndBeforeDateOffset.Hours) hours from today."
		Write-Host ""
		Write-Host -ForegroundColor Red "The EndBeforeDate of $($EndBeforeDate.ToShortDateString()) needs to be changed to today's date of" `
			"$($TodaysDate.ToShortDateString()) or prior, because today is not over and therefore doesn't comprise a required full 24 hour day."
		_Exit-Script -HardExit $True
	# Otherwise if the Start and End dates are the same days, report that and exit out of the script.
	} ElseIf ($StartOnDate -eq $EndBeforeDate) {
		Write-Host ""
		Write-Host -ForegroundColor Red "The StartOnDate and EndBeforeDate must be different dates, at least one full day apart."
		_Exit-Script -HardExit $True
	# Lastly if the Start date occurs after the End date, report that end exit out of the script.
	} ElseIf ($StartOnDate -gt $EndBeforeDate) {
		Write-Host ""
		Write-Host -ForegroundColor Red "The StartOnDate needs to be changed so it occurs before the EndBeforeDate."
		_Exit-Script -HardExit $True
	}
	Write-Verbose "The Start and End dates passed validation checks."

	# Validate if the minimum servers percentage per site is 100, so that all servers in a site are captured or the site is skipped, otherwise
	#   prompt the user to accept less than a complete data set per site.
	If ($MinServersPercent -lt 100) {
		Write-Host ""
		Write-Warning ("You have specified the `"MinServersPercent`" value of $MinServersPercent%, which instructs this script to " + `
			"tolerate $(100 - $MinServersPercent)% of the servers in a site being inaccessible/not returning data.")
		Write-Host "This could result in potentially skewed message profile results for a site, due to incomplete server data."
		# If Confirm wasn't set to False, then flush out any pending keys and then prompt to proceed.
		If ($Confirm) {
			Write-Host ""
			$Host.UI.RawUI.FlushInputBuffer()
			$Proceed = Read-Host "Do you want to proceed anyway? (Y/N)"
			If ($Proceed -eq "Y") {
				# Note this script will proceed.
				Write-Host "Continuing with script execution..."
			} Else {
				# Otherwise exit this script.
				_Exit-Script -HardExit $True
			}
		}
	}

	# Grab the number of logical cores on this system.
	$SystemCPUCores = (Get-WmiObject -Class Win32_Processor | Select-Object -ExpandProperty NumberOfCores | Measure-Object -Sum).Sum
	# Set the absolute maximum amount of threads to 12 since no Exchange server should have more than 24 cores and this script shouldn't place
	#   more than a 50% load on any production server. This also helps to limit the number of IIS based remote sessions to a host, so the
	#   associated session throttling limits shouldn't be exceeded (Exchange 2010's default session throttling limit is 18 per user).
	$AbsoluteMaxThreads = 12

	# If MaxThreads value was specified, and both it and SystemCPUCores have a value higher than the AbsoulteMaxThreads, then set MaxThreads
	#   to the AbsoluteMaxThreads value.
	If (($MaxThreads -gt $AbsoluteMaxThreads) -and ($SystemCPUCores -gt $AbsoluteMaxThreads)) {
		Write-Host ""
		Write-Warning "The specified `"MaxThreads`" value of $MaxThreads exceeds the absolute maximum value of $AbsoluteMaxThreads."
		Write-Host "Due to concerns with the target Exchange servers's IIS based remote shell throttling limits and overall CPU resource load," `
			"the value will be changed to $AbsoluteMaxThreads."
		$MaxThreads = $AbsoluteMaxThreads
	# Otherwise if the MaxThreads value was specified, and it has a higher value than SystemCPUCores, then set MaxThreas to that value.
	} ElseIf ($MaxThreads -gt $SystemCPUCores) {
		Write-Host ""
		Write-Warning "The specified `"MaxThreads`" value of $MaxThreads exceeds the $SystemCPUCores CPU cores on this system."
		Write-Host "To prevent overload of this system and from jobs competing with each other for resources, the value will be changed to" `
			"$SystemCPUCores."
		$MaxThreads = $SystemCPUCores
	}

	# Determine if the server is an Exchange server, and if so and calculate the recommended threads as 1/4 of the CPU cores to limit the CPU
	#   load, otherwise use the value of 1/2 of the CPU cores (both rounded up).
	If (Get-Service -DisplayName "*Microsoft Exchange*") {
		$RecommendedThreads = [Math]::Ceiling($SystemCPUCores / 4)
		$ThreadWarningText = "1/4 (rounded up) of this Exchange"
	} Else {
		$RecommendedThreads = [Math]::Ceiling($SystemCPUCores / 2)
		$ThreadWarningText = "1/2 (rounded up) of this"
	}

	# If MaxThreads was not specified as a parameter, then set it to the recommended threads.
	If (-Not($MaxThreads)) {
		$MaxThreads = $RecommendedThreads
	# Otherwise if Exchange is installed on this system and also if the MaxThreads exceeds the recommended threads, then write the warnings.
	} ElseIf ($MaxThreads -gt $RecommendedThreads) {
		Write-Host ""
		Write-Warning ("The specified `"MaxThreads`" value of $MaxThreads exceeds the recommended value of $RecommendedThreads, which is " + `
			"based on $ThreadWarningText server's $SystemCPUCores processor cores.")
		Write-Host "Running too many simultaneous jobs on this server could impact user performance by overloading CPU resources."
		# If Confirm wasn't set to False, then flush out any pending keys and then prompt to proceed.
		If ($Confirm) {
			Write-Host ""
			$Host.UI.RawUI.FlushInputBuffer()
			$Proceed = Read-Host "Do you want to proceed anyway? (Y/N)"
			If ($Proceed -eq "Y") {
				# Note this script will proceed.
				Write-Host "Continuing with script execution..."
			} Else {
				# Otherwise exit this script.
				_Exit-Script -HardExit $True
			}
		}
	}

	Write-Verbose "The maximum number simultaneous server data gathering jobs is $MaxThreads."
	#endregion ValidateInput

	#region GatherFunctions
	# Function to see if the percentage of skipped servers exceeds the MinServersPercent parameter based threshold.
	Function _Check-MinServersPercent {
		Param(
			[Parameter(Mandatory = $True, Position = 0)]
			[String]$SiteName,
			[Parameter(Mandatory = $True, Position = 1)]
			[Int]$ExchangeServerCount,
			[Parameter(Mandatory = $True, Position = 2)]
			[Int]$SkippedServerCount
		)

		# Calculate the remaining server percentage of servers in the site, rounding down, that haven't been skipped.
		$RemainingServerPercent = [Math]::Floor(100 - ($SkippedServerCount / $ExchangeServerCount * 100))

		# If the percentage of remaining servers in a site is below the parameter based threshold, report the site will be skipped and add its
		#   name to the SkippedSites array.
		If ($RemainingServerPercent -lt $MinServersPercent) {
			Write-Host ""
			# If there are some remaining servers then provide the more detailed explanation, otherwise just note all servers were inaccessible.
			If ($RemainingServerPercent -gt 0) {
				Write-Warning ("The script will now skip the site `"$SiteName`", including any remaining servers in it, and move on to the " +
					"next site in the collection.`nThis is because communication issues with $SkippedServerCount out of $ExchangeServerCount " +
					"servers exceeds the configuration that $MinServersPercent% of servers in a site respond with data.")
				Write-Host "To process this site anyway, re-run the script and adjust the -MinServersPercent to a number (without the `"%`"" `
					"sign) equal to or lower than the current $RemainingServerPercent% of servers accessible."
			} Else {
				Write-Warning "The script will now skip the site `"$SiteName`" because there were communication issues with all servers in it."
			}
			$Script:SkippedSites += $SiteName
			# If the site was listed in the PartialSites hash table then remove it since it will now be skipped.
			If ($PartialSites[$SiteName]) {
				$Script:PartialSites.Remove($SiteName)
			}
			# Shut down all running jobs.
			Get-Job | Where-Object {$_.Name -like "MessageProfile*"} | Remove-Job -Force -Confirm:$False
			# Continue on to the next site in the list.
			Continue ProcessSites
		# Otherwise if the SiteName exists in the PartialSites array, then update the recorded SkippedServers array value in it.
		} ElseIf ($PartialSites[$SiteName]) {
			$Script:PartialSites[$SiteName] = $SkippedServers
		# Otherwise add the SiteName to the PartialSites array with the SkippedServers array as the value.
		} Else {
			$Script:PartialSites.Add($SiteName,$SkippedServers)
		}
	}

	# ScriptBlock used as the "Script" in the data gathering job function directly below.
	$ServerScript = {
		Param ($ServerParameters)

		# Extract the parameters from the passed through ServerParameters, reconstituting the objects as necessary due to deserialization.
		$ExchangeServer = $ServerParameters.ExchangeServer
		$StartOnDate = Get-Date $ServerParameters.StartOnDate
		$EndBeforeDate = Get-Date $ServerParameters.EndBeforeDate
		$TimeSpan = $ServerParameters.TimeSpan
		$MessageFilter = $ServerParameters.MessageFilter
		$MailboxFilter = $ServerParameters.MailboxFilter

		# Extract the Server's FQDN and create a custom PS object to hold it and the server's gathered data.
		$ExchangeServerFQDN = $ExchangeServer.Fqdn
		$ServerData = New-Object PSCustomObject -Property @{
			ExchangeServerFQDN = $ExchangeServerFQDN
		}

		# Override the Write-Host and Write-Verbose cmdlets by turning them into empty functions so they won't cause any output to the pipeline.
		#   NOTE: The progress bar from Get-MessageTrackingLog cannot be suppressed unfortunately.
		Function Write-Host {}
		Function Write-Verbose {}

		# Dot source the remote exchange script to and then connect to an automatically chosen Exchange server so RBAC is enforced. Out-Null
		#   ensures any output not addressed by the overrides above is ignored.
		. $Env:ExchangeInstallPath\Bin\RemoteExchange.ps1 | Out-Null
		Connect-ExchangeServer -Auto

		#region GatherHubTransportData
		# If the server has the hub transport role, try to retrieve the message tracking log data from the current server using the specified
		#   start and end dates taking into account the time delta.
		If ($ExchangeServer.IsHubTransportServer -eq $True) {
			Try {
				# First retrieve the messages that were sent by a mailbox ("Received" by the Transport service from a mailbox via the
				#   STOREDRIVER).
				$SentMsgs = Get-MessageTrackingLog -Server $ExchangeServerFQDN -ResultSize:Unlimited `
					-Start $StartOnDate.AddHours($TimeSpan) -End $EndBeforeDate.AddHours($TimeSpan) -EventID Receive `
					-ErrorAction Stop | Where-Object ([ScriptBlock]::Create($MessageFilter)) | Select-Object Sender,TotalBytes
				# Next retrieve the messages that were received by a mailbox ("Delivered" by the Transport service to a
				#   mailbox via the STOREDRIVER).
				$ReceivedMsgs = Get-MessageTrackingLog -Server $ExchangeServerFQDN -ResultSize:Unlimited `
					-Start $StartOnDate.AddHours($TimeSpan) -End $EndBeforeDate.AddHours($TimeSpan) -EventID Deliver `
					-ErrorAction Stop | Where-Object ([ScriptBlock]::Create($MessageFilter)) | Select-Object Recipients,TotalBytes
			# If either message tracking log data retrieval failed, output the error to the debug stream, remove the remote PowerShell session,
			#   and return the Server Data object with the error added to it (ending any further job activity).
			} Catch {
				$ServerData | Add-Member -Type NoteProperty -Name ErrorMessage `
					-Value "Get-MessageTrackingLog at $((Get-Date).ToString()):`n $($_.Exception)"
				Get-PSSession | Remove-PSSession
				RETURN $ServerData
			}

			# Loop through each sent and received message on the server, adding the gathered per-server counts at the end of
			#   each loop.
			$ServerSentCount = 0
			$ServerSentSize = 0
			$ServerReceivedCount = 0
			$ServerReceivedSize = 0
			ForEach ($SentMsg in $SentMsgs) {
				# Increment the sent message count by 1 and add the message's size to the collection of sent message sizes.
				$ServerSentCount++
				$ServerSentSize += $SentMsg.TotalBytes
			}
			ForEach ($ReceivedMsg in $ReceivedMsgs) {
				# Because a received message could be delivered to multiple mailboxes in the same database, increment the
				#   received message count by the number of recipients the message, and add the message's size times the number
				#   of recipients to the collection of received message sizes.
				$RecipientCount = $ReceivedMsg.Recipients.Count
				$ServerReceivedCount += $RecipientCount
				$ServerReceivedSize += $ReceivedMsg.TotalBytes * $RecipientCount
			}

			# Add the collected message tracking data to the server data collection object.
			$ServerData | Add-Member -Type NoteProperty -Name TransportServer -Value $True
			$ServerData | Add-Member -Type NoteProperty -Name ServerSentCount -Value $ServerSentCount
			$ServerData | Add-Member -Type NoteProperty -Name ServerSentSize -Value $ServerSentSize
			$ServerData | Add-Member -Type NoteProperty -Name ServerReceivedCount -Value $ServerReceivedCount
			$ServerData | Add-Member -Type NoteProperty -Name ServerReceivedSize -Value $ServerReceivedSize
		}
		#endregion GatherHubTransportData

		#region GatherMailboxData
		# If the server has the mailbox role, set the AD scope to the entire forest and then try to gather the mailboxes on the Exchange server
		#   using the generated Mailbox filter.
		If ($ExchangeServer.IsMailboxServer -eq $True) {
			Set-ADServerSettings -ViewEntireForest $True | Out-Null
			Try {
				$Mailboxes = Get-Mailbox -Server $ExchangeServerFQDN -ResultSize:Unlimited `
					-Filter ([ScriptBlock]::Create($MailboxFilter)) -ErrorAction Stop | Select-Object PrimarySmtpAddress
			# If there was an error gathering the mailbox count, output the error to the debug stream, remove the remote PowerShell session, and
			# return the Server Data object with the error added to it (ending any further job activity).
			} Catch {
				$ServerData | Add-Member -Type NoteProperty -Name ErrorMessage `
					-Value "Get-Mailbox at $((Get-Date).ToString()):`n$($_.Exception)"
				Get-PSSession | Remove-PSSession
				RETURN $ServerData
			}

			# Add the collected mailbox data to the server data collection object.
			$ServerData | Add-Member -Type NoteProperty -Name MailboxServer -Value $True
			$ServerData | Add-Member -Type NoteProperty -Name ServerMailboxes -Value ($Mailboxes | Measure-Object).Count
		}
		#endregion GatherMailboxData

		# Record the number of PowerShell sessions to be verified when the job is retrieved.
		$ServerData | Add-Member -Type NoteProperty -Name PSSessions -Value (Get-PSSession | Measure-Object).Count

		# Remove the remote PowerShell session and return the collected server data to be extracted from the received job.
		Get-PSSession | Remove-PSSession
		RETURN $ServerData
	}

	# Function to add a data gathering job for the server.
	Function _Add-ServerJob {
		Param(
			[Parameter(Mandatory = $True, Position = 0)]
			$ExchangeServerFQDN
		)

		# If the server already exists in the ServerTries hash table, then it has already had one or more jobs run so increment the hash table
		#   entry for the server by 1.
		If ($ServerTries[$ExchangeServerFQDN]) {
			$Script:ServerTries[$ExchangeServerFQDN]++
		# Otherwise add the server to the hash table with the value of 1 to indicate this is the first try.
		} Else {
			$Script:ServerTries.Add($ExchangeServerFQDN,1)
		}

		# If the current number of tries has exceeded the defined max server of tries, report that no data could be retrieved from it, and add the
		#   server to the SkippedServers list.
		If ($ServerTries[$ExchangeServerFQDN] -gt $MaxServerTries) {
			Write-Host ""
			Write-Host -ForegroundColor Red "Unable to retrieve message profile data from $ExchangeServerFQDN. Skipping the server."
			$Script:SkippedServers += $ExchangeServerFQDN
			$Script:SkippedServerCount = ($SkippedServers | Measure-Object).Count
			# Also check to see if the MinServersPercent threshold has been crossed by the number of skipped servers, which will skip the site.
			_Check-MinServersPercent $SiteName $ExchangeServerCount $SkippedServerCount
			# If the check above passed, then increment the CurrentJobCount by one so the script won't think it needs to process the server again.
			$Script:CurrentJobCount++
		} Else {
			# Otherwise kick off another server data gathering job for it.
			Write-Verbose "+ Initiating data gathering attempt #$($ServerTries[$ExchangeServerFQDN]) for $ExchangeServerFQDN."
			# Splat the parameters to be passed into the job.
			$ServerParameters = @{
				ExchangeServer = $ExchangeServers | Where-Object {$_.Fqdn -eq $ExchangeServerFQDN}
				StartOnDate = $StartOnDate
				EndBeforeDate = $EndBeforeDate
				TimeSpan = $TimeSpan
				MessageFilter = $MessageFilter
				MailboxFilter = $MailboxFilter
			}
			# Start the data gathering job suppressing the output with [Void], and increment the CurrentJobCount by one.
			[Void](Start-Job -Name "MessageProfile-$ExchangeServerFQDN" -ScriptBlock $ServerScript -ArgumentList $ServerParameters)
			$Script:CurrentJobCount++
		}
	}
	#endregion GatherFunctions

	#region GatherADSites
	Write-Verbose "Gathering all AD sites that match the specified ADSite(s):`n$ADSites"
	# Create the Server Filter by looping through each specified AD site in the array, even if there is only one.
	$SiteArray = @()
	ForEach ($ADSite in $ADSites) {
		# Try to query for the specified site name (including those with wild cards).
		Try {
			$FoundSites = Get-ADSite $ADSite -ErrorAction Stop
		# If an error was encountered looking up the site name, report that and then exit out of the script.
		} Catch {
			Write-Host ""
			Write-Host -ForegroundColor Red "The AD site name $ADSite was not found. Please check your spelling and try again."
			_Exit-Script -HardExit $True
		}
		# Otherwise no error was encountered so loop through each site returned from the query (in case a wild card name was used).
		ForEach ($FoundSite in $FoundSites) {
			$SiteArray += $FoundSite.ToString()
		}
	}
	# Close out the $ServerFilter to include only Exchange 2007 or later servers that have Mailbox and/or HubTransport roles.
	$ServerFilter = '($SiteArray -Contains $_.Site) -and ($_.IsExchange2007OrLater -eq $True)'
	$ServerFilter += ' -and (($_.IsMailboxServer -eq $True) -or ($_.IsHubTransportServer -eq $True))'
	Write-Debug "The ServerFilter is:`n$ServerFilter"
	#endregion GatherADSites

	#region DynamicFilters
	# Set the initial message filter to include only messages that come from the Information Store.
	$MessageFilter = '($_.Source -eq "STOREDRIVER")'
	# Set the base Mailbox filter to exclude Equipment and Discovery mailboxes.
	$MailboxFilter = '(RecipientTypeDetails -ne "EquipmentMailbox") -and (RecipientTypeDetails -ne "DiscoveryMailbox")'
	# If the ExcludeHealthData switch was used, add HealthMailbox exclusions to the Message and Mailbox filters.
	If ($ExcludeHealthData) {
		$MessageFilter += ' -and ($_.Recipients -notlike "HealthMailbox*") -and ($_.Recipients -notlike "extest_*")'
		$MailboxFilter += ' -and (DisplayName -notlike "HealthMailbox*") -and (DisplayName -notlike "extest_*")'
	}
	# If the ExcludeJournalData switch was used, filter out all messages that end the MessageID with the text "@jounal.report.generator>",
	#   as those are always journal messages.
	If ($ExcludeJournalData) {
		$MessageFilter += ' -and ($_.MessageID -notlike "*journal.report.generator>")'
	}
	# If the ExcludePFData switch was used, add the PF messages subject line exclusions to the message filter.
	If ($ExcludePFData) {
		$MessageFilter += ' -and ($_.MessageSubject -ne "Folder Content") -and ($_.MessageSubject -notlike "*Backfill Response")'
		# NOTE: The following PF message subject lines are no included because users could possibly use them in day to day messages:
		#   "Backfill Request", "Status", and "Hierarchy".
	}
	# If the ExcludeRoomMailboxes switch was used, add the Room mailboxes to the exclusion filter.
	If ($ExcludeRoomMailboxes) {
		$MailboxFilter += ' -and (RecipientTypeDetails -ne "RoomMailbox")'
	}
	Write-Debug "The MessageFilter is:`n$MessageFilter"
	Write-Debug "The MailboxFilter is:`n$MailboxFilter"
	#endregion DynamicFilters

	#region GatherExchangeServers
	# Grab the current local computer time zone UTC offset.
	$LocalUTCOffset = ([System.TimeZoneInfo]::Local).BaseUTCOffset.Hours
	Write-Verbose "The local computer's UTC time zone offset is $LocalUTCOffset."
	# Determine the number of days between the start and end dates, and also how many days back the start date is.
	$TotalDays = ($EndBeforeDate - $StartOnDate).Days
	$StartDaysBack = ($TodaysDate - $StartOnDate).Days
	# Gather all of the Exchange servers in the specified AD site(s).
	Write-Host ""
	Write-Host -ForegroundColor Green "Gathering $TotalDays days worth of messaging activity across all defined Exchange server sites."
	Write-Host "In a short while a progress bar will appear."
	$ExchangeServers = Get-ExchangeServer * | Where-Object ([ScriptBlock]::Create($ServerFilter))
	# In case multiple sites were specified group all of the collected Exchange servers by the Exchange site they are in.
	$ExchangeSites = $ExchangeServers | Group-Object -Property Site -AsHashTable
	# Extract the number of AD sites with Exchange servers in them.
	$ExchangeSiteCount = $ExchangeSites.Count
	# Set the site loop count to 0 so it can be used to track the percentage of completion.
	$SiteLoopCount = 0
	# Array used to collect the name of any sites where no mailboxes or messages were found.
	$EmptySites = @()
	# Array used to collect the name of any sites that were skipped due to inaccessible Exchange servers or other errors.
	$SkippedSites = @()
	# Hash table used to collect the name of any sites that data from only some servers, and the count of missing servers.
	$PartialSites = @{}
	#endregion GatherExchangeServers

	# Loop through each Exchange site in the collection, sorting alphabetically by name, in a ForEach loop labeled "ProcessSites".
	:ProcessSites ForEach ($ExchangeSite in ($ExchangeSites.GetEnumerator() | Sort-Object -Property Name)) {
		# Array used to collect the names of servers that need data gathering jobs initiated.
		[System.Collections.ArrayList]$PendingJobs = @()
		# Array used to collect the names of the servers that have been skipped due to connectivity or data retrieval issues.
		$SkippedServers = @()
		# Hash table used to collect the names of servers the number of times they have run the data collection job so it can be compared to the
		#    MaxServerJobs parameter.
		$ServerTries = @{}
		# Set the TimeSpan variable to Null so that the remote time will be extracted below once for each new site that is processed.
		$TimeSpan = $Null
		# Set the following variables to 0 for every loop as initial starting values.
		[Int64]$SiteSentCount = 0
		[Int64]$SiteSentSize = 0
		[Int64]$SiteReceivedCount = 0
		[Int64]$SiteReceivedSize = 0
		$SiteMailboxes = 0
		$CompletedServerCount = 0
		$SkippedServerCount = 0

		# Retrieve the friendly name of the current Exchange Site. If it is one that should be excluded, then bypass processing it altogether by
		#   continuing on with the next site in the site collection, otherwise continue on.
		$SiteName = ($ExchangeSite.Name).Name
		If ($ExcludeSites -contains $SiteName) {
			Write-Verbose "***`"$SiteName`" is specified as an excluded site, so it will not be processed for data gathering."
			Continue ProcessSites
		} Else {
			Write-Verbose "+++ Beginning to process the Exchange site `"$SiteName`"."
		}

		# Calculate the percentage complete for the number of Exchange sites being processed, incrementing the SiteLoopCount by 1.
		$SitePercentComplete = [Math]::Round(($SiteLoopCount++ / $ExchangeSiteCount * 100),1)
		# Show a status bar while looping through all of the Exchange servers in the Exchange site.
		Write-Progress -Id 1 -Activity "Processing Exchange Servers in the AD Site: $SiteName" `
			-PercentComplete $SitePercentComplete -Status "$SitePercentComplete% Complete" `
			-CurrentOperation "Verifying connectivity to all Exchange servers in AD site #$SiteLoopCount out of $ExchangeSiteCount."

		# Extract the Exchange servers (sorting alphabetically by name) and their number in the current Exchange site.
		$ExchangeServers = $ExchangeSite.Value | Sort-Object -Property Name
		$ExchangeServerCount = ($ExchangeServers | Measure-Object).Count

		# Loop through all of the Exchange servers in a ForEach loop labeled "ProcessServers".
		$ServerLoopCount = 0
		:ProcessServers ForEach ($ExchangeServer in $ExchangeServers) {
			# Extract the Exchange Server's FQDN.
			$ExchangeServerFQDN = $ExchangeServer.Fqdn

			# Calculate the percentage complete for the number of Exchange servers processed, incrementing the ServerLoopCount by 1.
			$ServerPercentComplete = [Math]::Round(($ServerLoopCount++ / $ExchangeServerCount * 100),1)
			Write-Progress -Id 2 -ParentId 1 -Activity "Testing connectivity to the Exchange Server: $ExchangeServerFQDN" `
				-PercentComplete $ServerPercentComplete -Status "$ServerPercentComplete% Complete" `

			# Test to see if the server is reachable over the network.
			If (-Not(_Test-Connectivity $ExchangeServerFQDN)) {
				# The test failed so report that, add the server to the SkippedServers list, and check to see if the site should be skipped due
				#   to exceeding the MinServersPercent threshold.
				Write-Host -ForegroundColor Red "The connection test to the server $ExchangeServerFQDN failed."
				$SkippedServers += $ExchangeServerFQDN
				$SkippedServerCount = ($SkippedServers | Measure-Object).Count
				# Also check to see if the MinServersPercent threshold has been crossed by the number of skipped servers, which will skip the
				#   site.
				_Check-MinServersPercent $SiteName $ExchangeServerCount $SkippedServerCount
				# If the site wasn't skipped, then continue on to the next server in the list.
				Write-Host "Skipping $ExchangeServerFQDN and moving on to the next server in the list."
				Continue ProcessServers
			}

			#region TimeSpanCheck
			# If the TimeSpan variable does not currently have a value, then try to retrieve the date/time and UTF offset (dividing it by 60
			#   minutes to get the value in hours).
			If ($Null -eq $TimeSpan) {
				Write-Progress -Id 2 -ParentId 1 -Activity "Testing connectivity to the Exchange Server: $ExchangeServerFQDN" `
					-PercentComplete $ServerPercentComplete -Status "$ServerPercentComplete% Complete" `
					-CurrentOperation "Retrieving the time zone information for the site."
				Write-Verbose "+ Retrieving time zone information from $ExchangeServerFQDN."
				Try {
					$RemoteTime = Get-WmiObject -Class Win32_LocalTime -ComputerName $ExchangeServerFQDN -ErrorAction Stop | `
						ForEach-Object {Get-Date -Month $_.Month -Day $_.Day -Year $_.Year -Hour $_.Hour -Minute $_.Minute `
						-Second $_.Second}
					$RemoteUTCOffSet = ((Get-WmiObject -Class Win32_TimeZone -ComputerName $ExchangeServerFQDN `
						-ErrorAction Stop).Bias / 60)
				# If there was an error retrieving the remote computer time, report that and add the server to the SkippedServers list.
				} Catch {
					Write-Host ""
					Write-Host -ForegroundColor Red "Unable to retrieve the remote time information from $ExchangeServerFQDN."
					Write-Debug "The error returned at $((Get-Date).ToString()) was:`n$($_.Exception)"
					$SkippedServers += $ExchangeServerFQDN
					$SkippedServerCount = ($SkippedServers | Measure-Object).Count
					# Also check to see if the MinServersPercent threshold has been crossed by the number of skipped servers, which will skip
					#   the site.
					_Check-MinServersPercent $SiteName $ExchangeServerCount $SkippedServerCount
					# If the site wasn't skipped, then continue on to the next server in the list.
					Write-Host "Skipping $ExchangeServerFQDN and moving on to the next server in the list."
					Continue ProcessServers
				}
				# Retrieve the date and time on the local computer so it can be used to compare to the remote server below. This needs to be
				#   done every time the remote time is retrieved so the script compares current time and date stamps from both computers in case
				#   the script is taking a long time to run.
				$LocalTime = Get-Date
				# Extract the time span delta in hours between then local and remote computer in case the computers are in different time zones.
				#   This is important because the Get-MessageTrackingLog cmdlet -Start and -End parameters are always interpreted by the local
				#   computer, not the remote computer. Using TotalHours versus Hours, and then rounding it to 2 decimal point supports time spans
				#   that include 1/4 hour, and ensures scenarios where a delta of 2 hours and 59 minutes are recorded as 3 hours.
				$TimeSpan = [Math]::Round((New-TimeSpan -Start $LocalTime -End $RemoteTime).TotalHours,2)
				Write-Verbose ("*** The `"$SiteName`" site's time is $TimeSpan hours off from the local server, and has the UTC time " +
					"zone of $RemoteUTCOffSet.")
			}
			#endregion TimeSpanCheck

			# If the server was marked to be skipped, and just note it to the Verbose stream.
			If ($SkippedServers -contains $ExchangeServerFQDN) {
				Write-Verbose "* $ExchangeServerFQDN is marked to be skipped due to connectivity issues, so it will not be processed."

			# Otherwise perform the following actions.
			} Else {
				# If the server is registered as a HubTransport server, check if it's creation date is newer than the specified start date.
				If ($ExchangeServer.IsHubTransportServer -eq $True) {
					If ($ExchangeServer.WhenCreated	-gt $StartOnDate) {
						# It was so warn of the potential for incomplete message tracking data for the site.
						Write-Host ""
						Write-Warning ("$ExchangeServerFQDN was created after the specified StartOnDate of " + `
							"$($StartOnDate.ToShortDateString()). If this server replaced another server that was decommissioned " + `
							"after the StartOnDate, then some of the necessary message tracking data for the `"$SiteName`" site may be " + `
							"missing.")
					}
					# Next determine if the StartOnDatee is outside the server's number of days to keep log files (MessageTrackingLogAge), and
					#   if so report a warning that the specified date range is beyond this server's log retention.
					$TrackingLogAgeDays = (Get-TransportServer $ExchangeServerFQDN `
						-WarningAction SilentlyContinue).MessageTrackingLogMaxAge.Days
					If ($StartDaysBack -gt $TrackingLogAgeDays) {
						Write-Host ""
						Write-Warning ("$ExchangeServerFQDN is configured to only keep Message Tracking Logs for $TrackingLogAgeDays days, " + `
							"and the `"StartOnDate`" was set to $StartDaysBack days ago. This server will negatively skew the message " + `
							"profile by providing insufficient message tracking history for the intended number of days.")
						Write-Host "It is recommended to re-run this script for the site `"$SiteName`" with a `"StartOnDate`" value within" `
							"the last $TrackingLogAgeDays days."
					}
				}

				# Add it to the PendingJobs array, suppressing the output with [Void].
				[Void]$PendingJobs.Add($ExchangeServerFQDN)
			}
		}
		# Close out the Exchange server progress bar cleanly.
		Write-Progress -Id 2 -Completed -Activity "Testing connectivity to the Exchange Server:" -Status "Completed"

		# For job tracking purposes set the initial job count to the pending number of jobs, and set the current job count to 0.
		$InitialJobCount = $PendingJobs.Count
		$CurrentJobCount = 0

		# Update the primary progress bar by removing the comment about verifying server connectivity.
		Write-Progress -Id 1 -Activity "Processing Exchange Servers in the AD Site: $SiteName" -PercentComplete $SitePercentComplete `
			-Status "$SitePercentComplete% Complete" `
			-CurrentOperation "Tracking $InitialJobCount server jobs in AD site #$SiteLoopCount out of $ExchangeSiteCount."

		# Change the default behavior of CTRL-C so that the script can intercept and use it versus just terminating the script.
		[Console]::TreatControlCAsInput = $True
		# Sleep for 1 second and then flush the key buffer so any previously pressed keys are discarded and the loop can monitor for the use of
		#   CTRL-C. The sleep command ensures the buffer flushes correctly.
		Start-Sleep -Seconds 1
		$Host.UI.RawUI.FlushInputBuffer()

		# Continue to loop while there are pending or currently executing jobs.
		While ($PendingJobs -or $CurrentJobCount) {
			# If a key was pressed during the loop execution, check to see if it was CTRL-C (aka "3"), and if so exit the script after clearing
			#   out any running jobs and setting CTRL-C back to normal.
			If ($Host.UI.RawUI.KeyAvailable -and ($Key = $Host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp"))) {
				If ([Int]$Key.Character -eq 3) {
					Write-Host ""
					Write-Warning "CTRL-C was used - Shutting down any running jobs before exiting the script."
					Get-Job | Where-Object {$_.Name -like "MessageProfile*"} | Remove-Job -Force -Confirm:$False
					[Console]::TreatControlCAsInput = $False
					_Exit-Script -HardExit $True
				}
				# Flush the key buffer again for the next loop.
				$Host.UI.RawUI.FlushInputBuffer()
			}

			# Set ProgressUpdate to false so the progress bar at the end only updates if a job is added or removed (to save screen refresh time).
			$ProgressUpdate = $False

			# Loop through each Message Profile data gathering job.
			ForEach ($Job in (Get-Job | Where-Object {$_.Name -like "MessageProfile*"})) {
				# If the data gathering job is completed, then extract the results from it, remove it, decrement the current job count, and then
				#   mark for a progress bar update.
				If ($Job.State -eq "Completed") {
					$JobResult = Receive-Job $Job
					Remove-Job $Job
					$CurrentJobCount--
					$ProgressUpdate = $True

					# Extract the server FQDN from the job results.
					$ExchangeServerFQDN = $JobResult.ExchangeServerFQDN

					# If there was an error message then report it and re-run the job by re-adding it to the Pending Jobs array.
					If ($JobResult.ErrorMessage) {
						Write-Verbose ("* $ExchangeServerFQDN experienced an error during data gathering. Re-running the job up to " + `
							"$MaxServerTries times to ensure no data is missed.")
						Write-Debug "$ExchangeServerFQDN returned the following error for $($JobResult.ErrorMessage)"
						[Void]$PendingJobs.Add($ExchangeServerFQDN)
					# Next if there was more than one PowerShell session in the job, which means the original remote PowerShell session broke
					#   part of the way through and was re-established, re-run the job because it's possible there was some data loss/omission
					#  when the session broke.
					} ElseIf ([Int]$JobResult.PSSessions -gt 1 ) {
						Write-Verbose ("* $ExchangeServerFQDN had more then 1 PSSession which means the original session broke and some " + `
							"data loss could have occurred. Re-running the job up to $MaxServerTries to ensure no data is missed.")
						[Void]$PendingJobs.Add($ExchangeServerFQDN)
					} Else {
						# Otherwise the data gathering was successful so increment the Completed Servers count and process the results.
						$CompletedServerCount++
						Write-Verbose "> Processing the results from server $ExchangeServerFQDN."
						# If message tracking data was returned, then extract the values from the job results and then add them to the running
						#   totals for the site.
						If ($JobResult.TransportServer) {
							[Int64]$ServerSentCount = $JobResult.ServerSentCount
							[Int64]$ServerSentSize = $JobResult.ServerSentSize
							[Int64]$ServerReceivedCount = $JobResult.ServerReceivedCount
							[Int64]$ServerReceivedSize = $JobResult.ServerReceivedSize
							$SiteSentCount += $ServerSentCount
							$SiteSentSize += $ServerSentSize
							$SiteReceivedCount += $ServerReceivedCount
							$SiteReceivedSize += $ServerReceivedSize
							Write-Verbose ("There were $ServerSentCount sent and $ServerReceivedCount received messages on " + `
								"$ExchangeServerFQDN during the specified $TotalDays day(s).")
						}
						# If mailbox data was returned, then extract the value from the job results and then add it to the running total for
						#   the site.
						If ($JobResult.MailboxServer) {
							[Int]$ServerMailboxes = $JobResult.ServerMailboxes
							If ($ServerMailboxes -gt 0) {
								$SiteMailboxes += $ServerMailboxes
								Write-Verbose "There are $ServerMailboxes mailboxes on $ExchangeServerFQDN."
							} Else {
								Write-Verbose "There were no mailboxes found on $ExchangeServerFQDN."
							}
						}
						Write-Verbose "- Finished processing the results server $ExchangeServerFQDN."
					}
				}
			}

			# To ensure that the script doesn't try to add more jobs than are pending, get the updated count of pending and currently running
			#   server jobs, and set the JobThrottle to the smaller number between it and MaxThreads.
			$PendingJobCount = $PendingJobs.Count
			$CurrentJobCount = (Get-Job | Where-Object {$_.Name -like "MessageProfile*"} | Measure-Object).Count
			If (($PendingJobCount + $CurrentJobCount) -lt $MaxThreads) {
				$JobThrottle = $PendingJobCount + $CurrentJobCount
			} Else {
				$JobThrottle = $MaxThreads
			}

			# While the number of current jobs is less than the job throttle, add a job for the first server ([0]) in the Pending Jobs array,
			#   remove the server from the Pending Jobs array, and then mark for a progress bar update.
			While ($CurrentJobCount -lt $JobThrottle) {
				$ExchangeServerFQDN = $PendingJobs[0]
				_Add-ServerJob $ExchangeServerFQDN
				$PendingJobs.Remove($ExchangeServerFQDN)
				$ProgressUpdate = $True
			}

			# If there was a progress update, then update the progress bar with the new information.
			If ($ProgressUpdate) {
				# Calculate the remaining jobs from pending and current jobs, and write a progress bar with the percentage of completed jobs.
				$RemainingJobs = $PendingJobs.Count + $CurrentJobCount
				$JobsPercentComplete = [Math]::Round(($CompletedServerCount / $InitialJobCount * 100),1)
				Write-Progress -Id 2 -ParentId 1 -Activity "Processing $CurrentJobCount of the $RemainingJobs remaining server jobs." `
					-PercentComplete $JobsPercentComplete -Status "$JobsPercentComplete% Complete" `
					-CurrentOperation ("Completed Servers: $CompletedServerCount - Skipped Servers: $SkippedServerCount - " + `
						"Total Servers: $ExchangeServerCount")
			}

			# If there are pending server jobs or currently running jobs sleep for 5 seconds before running through the ServerJobs loop again.
			If ($PendingJobs -or $CurrentJobCount) {
				Start-Sleep -Seconds 5
			}
		}
		# Return CTRL-C back to its default behavior after the ServerJobs loop is done.
		[Console]::TreatControlCAsInput = $False

		# Close out the server progress bar cleanly.
		Write-Progress -Id 2 -Completed -Activity "Processing server data gathering jobs." -Status "Completed"

		#region VerifySiteData
		# If there were any mailboxes found in the site, record the name of the site as an empty site, and continue on with the next site in the
		#   site collection.
		If ($SiteMailboxes -eq 0) {
			Write-Host ""
			Write-Warning ("There were no mailboxes found in the `"$SiteName`" site, so the script will exclude it and move on to " + `
				"the next site in the collection.")
			$EmptySites += $SiteName
			Continue ProcessSites
		}
		# If there were no sent or received messages found in the site during the specified time frame, then record the name of the site as an
		#   empty site and continue with the next site in the collection.
		If (($SiteSentCount + $SiteReceivedCount) -eq 0) {
			Write-Host ""
			Write-Warning ("There were no messages sent or received in the `"$SiteName`" site during the specified time frame, so " + `
				"the script will exclude it and move on to the next site in the collection.")
			$EmptySites += $SiteName
			Continue ProcessSites
		}
		#endregion VerifySiteData

		# Add all of the collected data to the DataTable by calling the _Add-SiteData Function.
		_Add-SiteData $SiteName $SiteMailBoxes $SiteSentCount $SiteSentSize $SiteReceivedCount $SiteReceivedSize $RemoteUTCOffset `
			$TimeSpan $TotalDays

		# Close out the primary Exchange site progress bar cleanly.
		Write-Progress -Id 2 -Completed -Activity "Processing Exchange Servers in the AD Site:" -Status "Completed"
		Write-Verbose "--- Finished processing the Exchange site `"$SiteName`"."
	}
	Write-Verbose "Completed Exchange site data gathering at $((Get-Date).ToString())."

# Otherwise check to see if the Import ParameterSet was used, if so then try to import the CSV file.
} ElseIf ($PsCmdlet.ParameterSetName -like "Import") {
	Try {
		$ImportCSV = Import-CSV $InCSVFile
	# If an error was caught report it and exit out of the script so the issue with the file name/path can be fixed.
	} Catch {
		Write-Host ""
		Write-Host -ForegroundColor Red "There was an error importing the CSF file `"$InCSVFile`". Please check the file name and path" `
			"and try again."
		_Exit-Script -HardExit $True
	}
	Write-Verbose "Imported data from the existing CSV file: $InCSVFile"

	# Loop through each entry in the CSV file and add its data to the MessageProfile DataTable using the Add-SiteData function.
	ForEach ($CSVEntry in $ImportCSV) {
		# Extract the SiteName variable and check to see if it is one that should be excluded.
		$SiteName = $CSVEntry.SiteName
		If ($ExcludeSites -contains $SiteName) {
			# It is so bypass recording it into the DataTable.
			Write-Verbose "* `"$SiteName`" is specified as an excluded site so it will not imported into the DataTable."
		} Else {
			# It isn't so pass the variables from the CSV file to the _Add-SiteData function, converting them from string to their required
			#   format as necessary. Also multiple the two KB entries by 1KB so they are passed through as Bytes and not KB.
			_Add-SiteData $SiteName ([Int]$CSVEntry.Mailboxes) ([Int64]$CSVEntry.SentMsgs) ([Int64]$CSVEntry.SentKB * 1KB) `
				([Int64]$CSVEntry.RcvdMsgs) ([Int64]$CSVEntry.RcvdKB * 1KB) ([Double]$CSVEntry.UTCOffset) `
				([Double]$CSVEntry.TimeSpan) ([Int]$CSVEntry.TotalDays)
		}
	}

# Otherwise lastly if the Existing ParameterSet was used, check to make sure at least one site (row) was found in the DataTable, and exit out of
#   the script if it wasn't.
} ElseIf ($PsCmdlet.ParameterSetName -like "Existing") {
	Write-Verbose "Processing in-memory data only."
	If ($MessageProfile.Rows.Count -lt 1) {
		Write-Host ""
		Write-Host -ForegroundColor Red "There was no data in the `$MessageProfile table."
		_Exit-Script -HardExit $True
	}
}

# If the AverageAllSites parameter was used, loop through all the rows in the DataTable collecting the data into totals, excluding the
#   "~All Sites" row. Convert the two KB values to Bytes by multiplying by 1KB, for proper handling by the _Add-SiteData function.
If ($AverageAllSites) {
	Write-Host -ForegroundColor Green "Adding an `"~All Sites`" entry since the AverageAllSites parameter was used."
	ForEach ($ProfileRow in $MessageProfile) {
		If ($ProfileRow.SiteName -notlike "~All Sites") {
			$TotalMailboxes += $ProfileRow.Mailboxes
			$TotalSentCount += $ProfileRow.SentMsgs
			$TotalSentSize += $ProfileRow.SentKB * 1KB
			$TotalReceivedCount += $ProfileRow.RcvdMsgs
			$TotalReceivedSize += $ProfileRow.RcvdKB * 1KB
			$AllDays += $ProfileRow.TotalDays
			$SiteRows++
		}
	}
	# Aggregate the days by dividing the number of all days by the number of non-"~All Sites" rows in the DataTable.
	$AggregateDays = $AllDays / $SiteRows
	# Verify the aggregate days is a whole number to help reduce the chance that partial site results are skewing the summary
	#   average, and if so add all the collected data to DataTale with the average all sites name and UTC offset and Timespan values of 0.
	If ($AggregateDays.GetType().Name -like "Int32") {
		_Add-SiteData "~All Sites" $TotalMailboxes $TotalSentCount $TotalSentSize $TotalReceivedCount $TotalReceivedSize "0" "0" `
			$AggregateDays
	# Otherwise report an error to the screen indicating that there is a discrepancy in the collected data.
	} Else {
		Write-Host ""
		Write-Host -ForegroundColor Red "One or more sites has an inconsistent number of days as the rest of the sites. To avoid inaccurate" `
			"results the average message profile for all sites will not be calculated."
	}
}

#region Outputs
# If no sites (0 rows) made it into the DataTable, report that and do nothing else.
If ($MessageProfile.Rows.Count -eq 0) {
	Write-Host ""
	Write-Host -ForegroundColor Red "No site data was recorded so there is nothing to report/export."
# Otherwise if the OutCSVFile parameter was used, export the DataTable sorted by SiteName to the specified CSV and report the action.
} ElseIf ($OutCSVFile) {
	$MessageProfile | Sort-Object -Property SiteName | Export-CSV -NoTypeInformation $OutCSVFile
	$Global:MessageProfile = $Null
	Write-Host ""
	Write-Host -ForegroundColor Green "The collected data was exported to the `"$OutCSVFile`" CSV file."
# Otherwise save the $MessageProfile variable in the shell session so it can be further manipulated after the script finishes.
} Else {
	# Set the column to sort on to "SiteName".
	$MessageProfile.DefaultView.Sort = "SiteName"
	# Temporarily process the DataTable as a DataView to implement the SiteName column sorting, but convert it back to a DataTable.
	#   Also reset the primary key on the DataTable as it is lost during the temporary conversion.
	$MessageProfile = ($MessageProfile.DefaultView).ToTable()
	$MessageProfile.PrimaryKey = $MessageProfile.Columns["SiteName"]
	# Finally pass the sorted the Message Profile DataTable back to the shell as and report the action.
	$Global:MessageProfile = $MessageProfile
	Write-Host ""
	Write-Host -ForegroundColor Green "The collected data was saved to the `$MessageProfile table variable."
}

# If the Gather ParameterSet was used, then check the following collections that only populate during a gathering effort.
If ($PsCmdlet.ParameterSetName -like "Gather") {
	# If the EmptySites array has any entries in it, report them to the screen.
	If ($EmptySites) {
		Write-Host ""
		Write-Host "The following site(s) either had no mailboxes and/or messaging activity so they were not included in the report:"
		$EmptySites
	}

	# If the PartialSites hash table has an entries in it, report them to the screen with a warning.
	If ($PartialSites.Count -gt 0) {
		Write-Host ""
		Write-Warning "The following $($PartialSites.Count) sites had one or more servers skipped during data gathering:"
		$PartialSites | Format-Table @{Name="ADSite";Expression={$_.Name}},@{Name="Skipped Servers";Expression={@($_.Value)}} -AutoSize
		Write-Host "Subsequently the message profiles for those sites is only based on a partial data set. It is highly recommended to address" `
			"the connectivity issues to those servers, and then re-run this script against those sites to get a full message profile for them."
	}

	# If the SkippedSites array has any entries in it, report them to the screen with a warning.
	If ($SkippedSites) {
		Write-Host ""
		Write-Warning "The following site(s) were skipped due to data gathering issues:"
		$SkippedSites
		Write-Host ""
		Write-Host -ForegroundColor Cyan "To re-run this script focusing on the skipped sites, use the -ADSites switch with the" `
			"`$SkippedSites variable populated by this script."
		Write-Host 'Example: .\Generate-MessageProfile.ps1 -ADSites $SkippedSites -StartOnDate <start> -EndBeforeDate <end> ...'
		# Also pass the MissedSites variable back to the shell so it can be used for another optional script run targeting just skipped sites.
		$Global:SkippedSites = $SkippedSites
	}
}
#endRegion Outputs

# Exit the script gracefully via the _Exit-Script function defined above by not using the HardExit parameter.
_Exit-Script -HardExit $False
