<#
.NOTES
	Name: SizingChecker.ps1
	Author: Marc Nivens
	Requires: Exchange 2013 Management Shell, Microsoft Excel, and administrator rights on the target Exchange
	server as well as the local machine.
	Version History:
	1.0 - 3/30/2015 - Initial Public Release.
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
    Attempts to calculate the peak average CPU utilization on an Exchange 2013 Mailbox or 
    Mailbox/CAS server based on hardware, message profile, and current number of active/passive
    mailboxes.
.DESCRIPTION
    This script takes the current hardware configuration, user profile information from the 
    Exchange 2013 Sizing Calculator, and the current number of active and passive mailboxes 
    and attempts to give an expected average peak CPU utilization.  It uses the formulas in 
    the "Ask the Perf Guy: Sizing Exchange 2013 Deployment" blog located here:

    http://blogs.technet.com/b/exchange/archive/2013/05/06/ask-the-perf-guy-sizing-exchange-2013-deployments.aspx.

    The sizing calculator gives the estimated CPU utilization during a failure event.  This
    script will give you the estimated CPU utilization during normal peak operation.  Take
    the overall system processor utilization % during the busiest 4 hour window of the day and it 
    should be within 5-10% of the results of this script.  If CPU is running high but this
    script predicts it will run high, it is most likely a load or a sizing issue.  If not, you 
    may be dealing with another issue such as sudden or unexpected client activity.  You can also
    use this script to come up with an accurate mutliplier for higher than expected load from
    external or 3rd party products.

    If you are importing results from the sizing calculator you will need to run the script 
    from a workstation with the Exchange 2013 Management Shell and Microsoft Excel installed.

    If you are not importing data from the sizing calculator you may specify the values manually.  
    As with the calculator you can specify up to 4 Tiers or "Profiles".  If you specify more
    than one profile you will need to give the percentage of the total mailboxes that each 
    profile contains.  For example, say you have 5000 Profile 1 mailboxes and 15000 Profile 2
    mailboxes.  The percent of Profile 1 mailboxes would be 25 and Profile 2 would be 75.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified.  If it is not the 		
	local server is assumed.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local 		
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on 		
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER CalculatorFile
    Specify the Sizing Calculator file to load.  You may either load message profile and SPECint 
    rating information from an already filled out calculator file or enter the values manually.
.PARAMETER Verbose	
	This optional parameter enables verbose logging.
.PARAMETER MailboxProfile1
    Number of messages sent/recieved per day per mailboxes in increments of 50 for Profile 1.  
    Maps to the "Total Send/Receive Capability/Mailbox/Day" or Tier1SendReceive field in the sizing 
    calculator.  Default is 1 profile at 50 messages.
.PARAMETER MailboxProfile2
    Number of messages sent/recieved per day per mailboxes in increments of 50 for Profile 2.  
    Maps to the "Total Send/Receive Capability/Mailbox/Day" or Tier2SendReceive field in the sizing 
    calculator.
.PARAMETER MailboxProfile3
    Number of messages sent/recieved per day per mailboxes in increments of 50 for Profile 3.  
    Maps to the "Total Send/Receive Capability/Mailbox/Day" or Tier3SendReceive field in the sizing 
    calculator.
.PARAMETER MailboxProfile4
    Number of messages sent/recieved per day per mailboxes in increments of 50 for Profile 4.  
    Maps to the "Total Send/Receive Capability/Mailbox/Day" or Tier4SendReceive field in the sizing 
    calculator.
.PARAMETER Profile1Percentage
    Percent of total mailboxes that apply to Profile 1.  Default is 100.  You can specify
    up to 4 profiles, the total needs to equal 100.  Doesn't map directly to a field in
    the calculator, rather it takes the field "Total Number of Tier-1 User Mailboxes/
    Environment" number for each tier and gets the percentage of the total mailbox count
    each one applies to.  Default is 1 profile at 100%.
.PARAMETER Profile2Percentage
    Percent of total mailboxes that apply to Profile 2.  You can specify up to 4 profiles, 
    the total needs to equal 100.  Doesn't map directly to a field in the calculator, rather 
    it takes the field "Total Number of Tier-2 User Mailboxes/Environment" number for each tier and 
    gets the percentage of the total mailbox count each one applies to. 
.PARAMETER Profile3Percentage
    Percent of total mailboxes that apply to Profile 3.  You can specify up to 4 profiles, 
    the total needs to equal 100.  Doesn't map directly to a field in the calculator, rather 
    it takes the field "Total Number of Tier-3 User Mailboxes/Environment" number for each tier and 
    gets the percentage of the total mailbox count each one applies to. 
.PARAMETER Profile4Percentage
    Percent of total mailboxes that apply to Profile 4.  You can specify up to 4 profiles, 
    the total needs to equal 100.  Doesn't map directly to a field in the calculator, rather 
    it takes the field "Total Number of Tier-4 User Mailboxes/Environment" number for each tier and 
    gets the percentage of the total mailbox count each one applies to. 
.PARAMETER MailboxMultiplier1
    Multiplier to increase the active megacycle requirements per active mailbox footprint
    for mailboxes that require additional megacycles (for example, mailboxes that 
    utilize third party clients such as Unified Messaging or Mobile devices).
    Maps to the "Megacycles Multiplication Factor" or Tier1MCFactor value in the sizing 
    calculator.  Default is 1.0.
.PARAMETER MailboxMultiplier2
    Multiplier to increase the active megacycle requirements per active mailbox footprint
    for mailboxes that require additional megacycles (for example, mailboxes that 
    utilize third party clients such as Unified Messaging or Mobile devices).
    Maps to the "Megacycles Multiplication Factor" or Tier2MCFactor value in the sizing 
    calculator.  Default is 1.0.
.PARAMETER MailboxMultiplier3
    Multiplier to increase the active megacycle requirements per active mailbox footprint
    for mailboxes that require additional megacycles (for example, mailboxes that 
    utilize third party clients such as Unified Messaging or Mobile devices).
    Maps to the "Megacycles Multiplication Factor" or Tier3MCFactor value in the sizing 
    calculator.  Default is 1.0.
.PARAMETER MailboxMultiplier4
    Multiplier to increase the active megacycle requirements per active mailbox footprint
    for mailboxes that require additional megacycles (for example, mailboxes that 
    utilize third party clients such as Unified Messaging or Mobile devices).
    Maps to the "Megacycles Multiplication Factor" or Tier4MCFactor value in the sizing 
    calculator.  Default is 1.0.
.PARAMETER Profile1MultiplierPercentage
    Percentage of users in the profile that will have the multiplier applied.  Maps
    to the "Multiplication factor user percentage" or Tier1UserMFPercentage field in 
    the sizing calculator.  Default is 100.
.PARAMETER Profile2MultiplierPercentage
    Percentage of users in the profile that will have the multiplier applied.  Maps
    to the "Multiplication factor user percentage" or Tier2UserMFPercentage field in 
    the sizing calculator.  Default is 100.
.PARAMETER Profile3MultiplierPercentage
    Percentage of users in the profile that will have the multiplier applied.  Maps
    to the "Multiplication factor user percentage" or Tier3UserMFPercentage field in 
    the sizing calculator.  Default is 100.
.PARAMETER Profile4MultiplierPercentage
    Percentage of users in the profile that will have the multiplier applied.  Maps
    to the "Multiplication factor user percentage" or Tier4UserMFPercentage field in 
    the sizing calculator.  Default is 100.
.PARAMETER SPECRating
    The Per-Core SPEC rating for the server.  This is the total SPECint 2006 rating for 
    the system divided by the number of physical cores.  Maps to the "Primary Datacenter
    Mailbox Servers\SPECint 2006 Rate Value" or numSpecRatePDC/numSpecRatePDC field
    in the sizing calculator, divided by the number of cores.  Default is 35.83.  For
    an accurate result please use the sizing blog to determine your exact SPECint 2006 
    rating, or get it from the sizing calculator.
.PARAMETER SecondaryDataCenter
    If importing numbers from the sizing calculator, use the processor/SPEC information from 
    the secondary datacenter section.
.EXAMPLE
	.\SizingChecker.ps1 -CalculatorFile E:\Files\calculator.xlsm -Server SERVERNAME -Verbose

	Run against a single remote Exchange server using calculator results with verbose logging.
.EXAMPLE
    .\SizingChecker.ps1 -Server SERVERNAME -MailboxProfile1 50 -MailboxProfile2 150 -Profile1Percentage 25 -Profile2Percentage 75 -MailboxMultiplier1 1.0 -MailboxMultiplier2 1.2 -Profile1MultiplierPercentage 100 -Profile2MultiplierPercentage 35 -SPECRating 41.5625

    The example above is based on two profiles.
      
    System SPECint 2006 rating: 665
    Number of physical cores: 16
    Per-Core SPEC rating (total SPEC rating/physical cores): 41.5625

    Profile 1:  
    Messages/day: 50 messages
    Mailboxes: 5000
    Multiplier 1.0
    % of mailboxes to apply multiplier to: 100

    Profile 2:  
    Messages/day: 150 messages
    Mailboxes: 15000
    Multiplier: 1.2
    % of mailboxes to apply multiplier to: 35 (35 percent of the mailboxes in this profile
    get a multiplier of 1.2, the rest get no multiplier or 1.0)
.EXAMPLE
	.\SizingChecker.ps1 -CalculatorFile E:\Files\calculator.xlsm -Server SERVERNAME -SecondaryDataCenter
	Run against a single remote Exchange server using the SPECint rating for the secondary datacenter.
.LINK
    http://blogs.technet.com/b/exchange/archive/2013/05/06/ask-the-perf-guy-sizing-exchange-2013-deployments.aspx.
#>

# Use the CmdletBinding function so the script accepts and understands -Verbose and -Debug and sets the default parameter set to
# "Gather". The Write-Verbose and Write-Debug statements in this script will activate only if their respective switches are used.
[CmdletBinding(DefaultParameterSetName = "Gather")]

#Parameters
param
(
    $Server = ($env:COMPUTERNAME), 
    $CalculatorFile = $null,
    $SPECRating = 35.83, 
    #[ValidateSet(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500)]
    $MailboxProfile1 = 50, 
    $MailboxProfile2 = 0, 
    $MailboxProfile3 = 0, 
    $MailboxProfile4 = 0, 
    $MailboxMultiplier1 = 1, 
    $MailboxMultiplier2 = 1, 
    $MailboxMultiplier3 = 1, 
    $MailboxMultiplier4 = 1, 
    [ValidateRange(0,100)]$Profile1Percentage = 100, 
    [ValidateRange(0,100)]$Profile2Percentage = 0, 
    [ValidateRange(0,100)]$Profile3Percentage = 0, 
    [ValidateRange(0,100)]$Profile4Percentage = 0, 
    [ValidateRange(0,100)]$Profile1MultiplierPercentage = 100, 
    [ValidateRange(0,100)]$Profile2MultiplierPercentage = 0, 
    [ValidateRange(0,100)]$Profile3MultiplierPercentage = 0, 
    [ValidateRange(0,100)]$Profile4MultiplierPercentage = 0,
    [ValidateScript({-not $_.ToString().EndsWith('\')})]$OutputFilePath = ".", 
    [switch]$SecondaryDataCenter
)

# Check to see if the -Verbose parameter was used.
If ($PSBoundParameters["Verbose"]) {
	#Write verbose output in Cyan since we already use yellow for warnings
	$VerboseForeground = $Host.PrivateData.VerboseForegroundColor
	$Host.PrivateData.VerboseForegroundColor = "Cyan"
}

#Enums and custom data types.  Use a trap here to handle the error if the types already exist (other health checker scripts have already added them)
Add-Type -TypeDefinition @"
    namespace SizingChecker
    {
        public enum ServerRole
        {
            MultiRole,
            Mailbox,
            ClientAccess,
            None
        }
        public enum ServerType
        {
            VMWare,
            HyperV,
            Physical,
            Unknown
        }
    }
"@

#Script Globals
#Versioning
$ScriptName = "Exchange 2013 Sizing Checker"
$ScriptVersion = "1.0"
$OutputFileName = "SizingCheck" + "-" + $Server + "-" + (get-date).tostring("MMddyyyyHHmmss") + ".log"
$OutputFullPath = $OutputFilePath + "\" + $OutputFileName

#System Information
$proc = Get-WmiObject -ComputerName $Server -Class Win32_Processor
$os = Get-WmiObject -ComputerName $Server -Class Win32_OperatingSystem
$system = Get-WmiObject -ComputerName $Server -Class Win32_ComputerSystem

#Megacycle Calculations
$script:P1ActiveMegacycles = 2.93
$script:P1PassiveMegacycles = .69
$script:P2ActiveMegacycles = 0
$script:P2PassiveMegacycles = 0
$script:P3ActiveMegacycles = 0
$script:P3PassiveMegacycles = .0
$script:P4ActiveMegacycles = 0
$script:P5PassiveMegacycles = 0

#Message, multipliers, and percentages
$script:P1Messages = $MailboxProfile1
$script:P2Messages = $MailboxProfile2
$script:P3Messages = $MailboxProfile3
$script:P4Messages = $MailboxProfile4
$script:P1Multiplier = $MailboxMultiplier1
$script:P2Multiplier = $MailboxMultiplier2
$script:P3Multiplier = $MailboxMultiplier3
$script:P4Multiplier = $MailboxMultiplier4
$script:P1Percent = $Profile1Percentage
$script:P2Percent = $Profile2Percentage
$script:P3Percent = $Profile3Percentage
$script:P4Percent = $Profile4Percentage
$script:P1MultiplierPercent = $Profile1MultiplierPercentage
$script:P2MultiplierPercent = $Profile2MultiplierPercentage
$script:P3MultiplierPercent = $Profile3MultiplierPercentage
$script:P4MultiplierPercent = $Profile4MultiplierPercentage

#SPEC rating
$script:PAllSpecRating = $SPECRating

#Processor Information
$script:NumberOfCores = $null
$script:NumberOfLogicalProcessors = $null
$script:MegacyclesPerCore = $null
$script:ProcessorName = $null


##################
#Helper Functions#
##################

#Output functions
function Write-Red($message)
{
    Write-Host $message -ForegroundColor Red
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Yellow($message)
{
    Write-Host $message -ForegroundColor Yellow
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Green($message)
{
    Write-Host $message -ForegroundColor Green
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Grey($message)
{
    Write-Host $message
    $message | Out-File ($OutputFullPath) -Append
}

function Write-VerboseOutput($message)
{
    Write-Verbose $message
    if($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
    {
        $message | Out-File ($OutputFullPath) -Append
    }
}

#Populate Needed active megacycles based on messages per day
#mailbox only is about .0425 per message/day
#AIO is about .0585
function Get-ActiveMailboxRequiredMegacycles($MailboxProfile)
{
    Write-VerboseOutput("Calling Get-ActiveMailboxRequiredMegacycles")

    if($ServerRole -eq [SizingChecker.ServerRole]::Mailbox)
    {
        switch($MailboxProfile)
        {
            50 {2.13}
            100 {4.25}
            150 {6.38}
            200 {8.50}
            250 {10.63}
            300 {12.75}
            350 {14.88}
            400 {17.00}
            450 {19.13}
            500 {21.25}
            default {[math]::Round(($MailboxProfile * .0425), 2)}
        }
    }
    elseif($ServerRole -eq [SizingChecker.ServerRole]::MultiRole)
    {
        switch($MailboxProfile)
        {
            50 {2.93}
            100 {5.84}
            150 {8.77}
            200 {11.69}
            250 {14.62}
            300 {17.53}
            350 {20.46}
            400 {23.38}
            450 {26.30}
            500 {29.22}
            default {[math]::Round(($MailboxProfile * .0585), 2)}
        }
    }
}

#Populate Needed passive megacycles based on messages per day
#about .0137 megacycles per message/day
function Get-PassiveMailboxRequiredMegacycles($MailboxProfile)
{
    Write-VerboseOutput("Calling Get-PassiveMailboxRequiredMegacycles")

    switch($MailboxProfile)
    {
        50 {0.69}
        100 {1.37}
        150 {2.06}
        200 {2.74}
        250 {3.43}
        300 {4.11}
        350 {4.80}
        400 {5.48}
        450 {6.17}
        500 {6.85}
        default {[math]::Round(($MailboxProfile * .0137), 2)}
    }
}

#Check Server Role (Mailbox, CAS, Both)
function Get-ServerRole
{
    Write-VerboseOutput("Calling Get-ServerRole")

    $role = (Get-ExchangeServer $Server).ServerRole
    if($role -eq "Mailbox, ClientAccess")
    {
        [SizingChecker.ServerRole]::MultiRole
        return
    }
    elseif($role -eq "Mailbox")
    {
        [SizingChecker.ServerRole]::Mailbox
        return
    }
    elseif($role -eq "ClientAccess")
    {
        [SizingChecker.ServerRole]::ClientAccess
        return
    }
    else
    {
        [SizingChecker.ServerRole]::None
        return
    }
}

#Check .NET Framework Version
#REMOVED

#Write All Reported Information to the console
function Write-SystemInformationToConsole
{
    Write-VerboseOutput("Calling Write-SystemInformationToConsole")

    Write-Green($ScriptName + " version " + $ScriptVersion)
    Write-Green("Sizing Information Report for " + $Server + " on " + (Get-Date) + "`r`n")

    #OS Version
    Write-Grey("Operating System: " + $os.Caption)

    #Exchange Version
    Write-Grey("Exchange: " + (Get-ExchangeServer $Server).AdminDisplayVersion)

    #Virtualized
    if($system.Manufacturer -like "VMWare*")
    {
        $ServerType = [SizingChecker.ServerType]::VMWare
    }
    elseif($system.Manufacturer -like "Microsoft Corporation")
    {
        $ServerType = [SizingChecker.ServerType]::HyperV
    }
    elseif($system.Manufacturer.Length -gt 0)
    {
        $ServerType = [SizingChecker.ServerType]::Physical
    }
    else
    {
        $ServerType = [SizingChecker.ServerType]::Unknown
    }
    Write-Grey("Hardware Type: " + $ServerType.ToString())

    #ServerRole
    Write-Grey("Server Role: " + $ServerRole.ToString())
    
    #Report Processor Information
    Write-VerboseOutput("Processor/Memory Information")
    Write-VerboseOutput("`tProcessor Type: " + $script:ProcessorName)
    Write-VerboseOutput("`tNumber of Processors: " + $system.NumberOfProcessors)
    Write-VerboseOutput("`tNumber of Physical Cores: " + $script:NumberOfCores)
    Write-VerboseOutput("`tNumber of Logical Cores: " + $script:NumberOfLogicalProcessors)
    Write-VerboseOutput("`tMegacycles Per Core: "  + $script:MegacyclesPerCore)
    if($CalculatorFile -eq $null)
    {
        Write-VerboseOutput("`Per-Core SPECint 2006 Rating: " + $script:PAllSpecRating)
    }
}

#Check if the server has the necessary processing power and amount of memory given the number of mailboxes
function Check-NeededMegacycles
{
    Write-VerboseOutput("Calling Check-NeededMegacycles")

    #Available megacycles using server SPECInt 2006 rating
    if($script:NumberOfCores -gt 0)
    {
        $AvailableMegacycles = [math]::Round((($script:PAllSpecRating * 2000)/33.75) * $script:NumberOfCores, 0)
    }
    else
    {
        Write-Red("Number of cores detected is zero.  Cannot calculate available megacycles.")
        Exit
    }

    
    Write-VerboseOutput("Profile 1 Megacycles per active mailbox needed: " + $script:P1ActiveMegacycles)
    Write-VerboseOutput("Profile 1 Megacycles per passive mailbox needed: " + $script:P1PassiveMegacycles)
    Write-VerboseOutput("Profile 1 Megacycles Multiplication Factor: " + $script:P1Multiplier)
    Write-VerboseOutput("Profile 1 Megacycles Multiplier user percentage: " + $script:P1MultiplierPercent)

    if($script:P2Percent -ne 0)
    {
        Write-VerboseOutput("Profile 2 Megacycles per active mailbox needed: " + $script:P2ActiveMegacycles)
        Write-VerboseOutput("Profile 2 Megacycles per passive mailbox needed: " + $script:P2PassiveMegacycles)
        Write-VerboseOutput("Profile 2 Megacycles Multiplication Factor: " + $script:P2Multiplier)
        Write-VerboseOutput("Profile 2 Megacycles Multiplier user percentage: " + $script:P2MultiplierPercent)
    }

    if($script:P3Percent -ne 0)
    {
        Write-VerboseOutput("Profile 3 Megacycles per active mailbox needed: " + $script:P3ActiveMegacycles)
        Write-VerboseOutput("Profile 3 Megacycles per passive mailbox needed: " + $script:P3PassiveMegacycles)
        Write-VerboseOutput("Profile 3 Megacycles Multiplication Factor: " + $script:P3Multiplier)
        Write-VerboseOutput("Profile 3 Megacycles Multiplier user percentage: " + $script:P3MultiplierPercent)
    }

    if($script:P4Percent-ne 0)
    {
        Write-VerboseOutput("Profile 4 Megacycles per active mailbox needed: " + $script:P4ActiveMegacycles)
        Write-VerboseOutput("Profile 4 Megacycles per passive mailbox needed: " + $script:P4PassiveMegacycles)
        Write-VerboseOutput("Profile 4 Megacycles Multiplication Factor: " + $script:P4Multiplier)
        Write-VerboseOutput("Profile 4 Megacycles Multiplier user percentage: " + $script:P4MultiplierPercent)
    }
    
    Write-Grey("Calculating needed CPU based on active and passive mailbox total")

    #Get a list of all active database copies on this server and total the number of mailboxes on each DB
    $MountedDBs = Get-MailboxDatabaseCopyStatus -server $Server | ?{$_.Status -eq 'Mounted'}
    if($MountedDBs.Count -gt 0)
    {
        $MountedDBs.DatabaseName | %{Write-VerboseOutput "Calculating Mailbox Total for Active Database: $_";$TotalActiveMailboxCount+=(get-mailbox -Database $_ -ResultSize Unlimited).Count}
        Write-VerboseOutput("Total Active Mailboxes on server: " + $TotalActiveMailboxCount)
    }
    else
    {
        Write-Warning "No Active Mailboxes found on server."
    }

    #Get a list of all passive healthy database copies on this server and total the number of mailboxes on each DB
    $PassiveDBs = Get-MailboxDatabaseCopyStatus -server $Server | ?{$_.Status -eq 'Healthy'}
    if($PassiveDBs.Count -gt 0)
    {
        $PassiveDBs.DatabaseName | %{Write-VerboseOutput "Calculating Mailbox Total for Passive Database: $_";$TotalPassiveMailboxCount+=(get-mailbox -Database $_ -ResultSize Unlimited).Count}
        Write-VerboseOutput("Total Passive Mailboxes on server: " + $TotalPassiveMailboxCount)
    }
    else
    {
        Write-VerboseOutput "No Passive Mailboxes found on server."
    }

    if(($PassiveDBs.Count -eq 0) -and ($MountedDBs.Count -eq 0))
    {
        Write-Warning "No Database Copies detected on server."
        break
    }

    #Calculate the needed megacycles.  The formula is documented here: http://blogs.technet.com/b/exchange/archive/2013/05/06/ask-the-perf-guy-sizing-exchange-2013-deployments.aspx
    #The formula for available megacycles is ((Platform Per Core SPEC rating x 2000)/37.5) x number of cores.  We use the baseline platorm score to estimate CPU needed if the SPEC rating for the server
    #is not passed in at script execution.  Since the sizing calculator allows for multiple usage profiles we have to account for that as well.  Say 80% of the users are 50 messages/day
    #with a multiplier of 1, and 20% are 100/day with a multiplier of 1.5.  We allow up to 4 profiles to be passed in to the script.  For each profile we do this:
    # 1. Get the total active mailbox count (using percentage for the profile)
    # 2. Get the total passive mailbox count (using percentage for the profile)
    # 3. Separate the mailboxes the multiplier is to be applied to vs. the ones it is not to be applied to 
    # 4. Multiply total active mailbox count by the required megacycles according to messages/day.  Multiply the result by the mailbox multiplier for percentage of mailboxes that need it.
    #    Those that do not need it, just multiply by required megacycles.  Total the two numbers.
    # 5. Multiply total passive mailbox count by the required megacycles according to messages/day.  Mailbox multiplier does not apply to passive mailboxes.
    # 6. Total the sums from 4 and 5 to get total required megacycles for the profile.
    # 7. Once all 4 profiles are done, total the required megacycles for each and check against the available megacycles for the server.
    $P1ActiveMailboxTotal = [math]::Round(($TotalActiveMailboxCount * ($script:P1Percent * .01)), 0)
    Write-VerboseOutput("Profile 1 Active Mailbox total: " + $P1ActiveMailboxTotal)
    $P1PassiveMailboxTotal = [math]::Round(($TotalPassiveMailboxCount * ($script:P1Percent * .01)), 0)
    Write-VerboseOutput("Profile 1 Passive Mailbox total: " + $P1PassiveMailboxTotal)
    $P1MultiplierMailboxes = [math]::Round($P1ActiveMailboxTotal * ($script:P1MultiplierPercent * .01), 0)
    Write-VerboseOutput("Multiplier applied to " + $P1MultiplierMailboxes + " mailboxes")
    $P1NonMultiplierMailboxes = [math]::Round($P1ActiveMailboxTotal - $P1MultiplierMailboxes, 0)
    Write-VerboseOutput("Multiplier not applied to " + $P1NonMultiplierMailboxes + " mailboxes")
    $P1NeededMegacycles = [math]::Round((($P1MultiplierMailboxes * ($script:P1ActiveMegacycles * $script:P1Multiplier)) + ($P1NonMultiplierMailboxes * $script:P1ActiveMegacycles)) + ($script:P1PassiveMegacycles * $P1PassiveMailboxTotal), 0)
    Write-VerboseOutput("Megacycles needed for profile 1: " + $P1NeededMegacycles)

    if($script:P2Percent -ne 0)
    {
        $P2ActiveMailboxTotal = [math]::Round(($TotalActiveMailboxCount * ($script:P2Percent * .01)), 0)
        Write-VerboseOutput("Profile 2 Active Mailbox total: " + $P2ActiveMailboxTotal)
        $P2PassiveMailboxTotal = [math]::Round(($TotalPassiveMailboxCount * ($script:P2Percent * .01)), 0)
        Write-VerboseOutput("Profile 2 Passive Mailbox total: " + $P2PassiveMailboxTotal)
        $P2MultiplierMailboxes = [math]::Round($P2ActiveMailboxTotal * ($script:P2MultiplierPercent * .01), 0)
        Write-VerboseOutput("Multiplier applied to " + $P2MultiplierMailboxes + " mailboxes")
        $P2NonMultiplierMailboxes = [math]::Round($P2ActiveMailboxTotal - $P2MultiplierMailboxes, 0)
        Write-VerboseOutput("Multiplier not applied to " + $P2NonMultiplierMailboxes + " mailboxes")
        $P2NeededMegacycles = [math]::Round((($P2MultiplierMailboxes * ($script:P2ActiveMegacycles * $script:P2Multiplier)) + ($P2NonMultiplierMailboxes * $script:P2ActiveMegacycles)) + ($script:P2PassiveMegacycles * $P2PassiveMailboxTotal), 0)
        Write-VerboseOutput("Megacycles needed for profile 2: " + $P2NeededMegacycles)
    }
    if($script:P3Percent -ne 0)
    {
        $P3ActiveMailboxTotal = [math]::Round(($TotalActiveMailboxCount * ($script:P3Percent * .01)), 0)
        Write-VerboseOutput("Profile 3 Active Mailbox total: " + $P3ActiveMailboxTotal)
        $P3PassiveMailboxTotal = [math]::Round(($TotalPassiveMailboxCount * ($script:P3Percent * .01)), 0)
        Write-VerboseOutput("Profile 3 Passive Mailbox total: " + $P3PassiveMailboxTotal)
        $P3MultiplierMailboxes = [math]::Round($P3ActiveMailboxTotal * ($script:P3MultiplierPercent * .01), 0)
        Write-VerboseOutput("Multiplier applied to " + $P3MultiplierMailboxes + " mailboxes")
        $P3NonMultiplierMailboxes = [math]::Round($P3ActiveMailboxTotal - $P3MultiplierMailboxes, 0)
        Write-VerboseOutput("Multiplier not applied to " + $P3NonMultiplierMailboxes + " mailboxes")
        $P3NeededMegacycles = [math]::Round((($P3MultiplierMailboxes * ($script:P3ActiveMegacycles * $script:P3Multiplier)) + ($P3NonMultiplierMailboxes * $script:P3ActiveMegacycles)) + ($script:P3PassiveMegacycles * $P3PassiveMailboxTotal), 0)
        Write-VerboseOutput("Megacycles needed for profile 3: " + $P3NeededMegacycles)
    }
    if($script:P4Percent -ne 0)
    {
        $P4ActiveMailboxTotal = [math]::Round(($TotalActiveMailboxCount * ($script:P4Percent * .01)), 0)
        Write-VerboseOutput("Profile 4 Active Mailbox total: " + $P4ActiveMailboxTotal)
        $P4PassiveMailboxTotal = [math]::Round(($TotalPassiveMailboxCount * ($script:P4Percent * .01)), 0)
        Write-VerboseOutput("Profile 4 Passive Mailbox total: " + $P4PassiveMailboxTotal)
        $P4MultiplierMailboxes = [math]::Round($P4ActiveMailboxTotal * ($script:P4MultiplierPercent * .01), 0)
        Write-VerboseOutput("Multiplier applied to " + $P4MultiplierMailboxes + " mailboxes")
        $P4NonMultiplierMailboxes = [math]::Round($P4ActiveMailboxTotal - $P4MultiplierMailboxes, 0)
        Write-VerboseOutput("Multiplier not applied to " + $P4NonMultiplierMailboxes + " mailboxes")
        $P4NeededMegacycles = [math]::Round((($P4MultiplierMailboxes * ($script:P4ActiveMegacycles * $script:P4Multiplier)) + ($P4NonMultiplierMailboxes * $script:P4ActiveMegacycles)) + ($script:P4PassiveMegacycles * $P4PassiveMailboxTotal), 0)
        Write-VerboseOutput("Megacycles needed for profile 4: " + $P4NeededMegacycles)
    }

    #Total megacycles needed from all 4 profiles
    $TotalNeededMegacycles = ($P1NeededMegacycles + $P2NeededMegacycles + $P3NeededMegacycles + $P4NeededMegacycles)
    
    #Estimated CPU is just needed megacycles divided by available megacycles
    $EstimatedCPUUsage = "{0:P0}" -f ($TotalNeededMegacycles/$AvailableMegacycles)
    Write-Grey("Total available megacycles: " + $AvailableMegacycles)
    Write-Grey("Total needed megacycles: " + $TotalNeededMegacycles)
    if($EstimatedCPUUsage -le 50)
    {
        Write-Green("Estimated Average CPU Usage: " + $EstimatedCPUUsage)
    }
    elseif($EstimatedCPUUsage -le 75)
    {
        Write-Yellow("Estimated Average CPU Usage: " + $EstimatedCPUUsage)
    }
    else
    {
        Write-Red("Estimated Average CPU Usage: " + $EstimatedCPUUsage)
    }
}

function Set-ProfileValuesFromSpreadsheet
{
    Write-VerboseOutput("Calling Set-ProfileValuesFromSpreadsheet")

    #see if we can get to the file
    if((Test-Path $CalculatorFile) -eq $false)
    {
        Write-Red("Invalid calculator path or file name: " + $CalculatorFile)
    }

    Write-VerboseOutput("Opening Sizing Calculator Results file " + $CalculatorFile)

    #Open spreadsheet as an excel object
    $Spreadsheet = New-Object -ComObject Excel.Application
    if($Spreadsheet -eq $null)
    {
        Write-Red("To read values from the calculator spreadsheet you must run this script from a machine that has Excel installed.  Please run this script from a workstation with the Exchange 2013 management tools and Microsoft Excel installed.")
        Exit
    }

    $Spreadsheet.Visible = $false
    $Workbook = $Spreadsheet.Workbooks.Open($CalculatorFile)
    $InputSheet = $Workbook.worksheets | where {$_.name -eq "Input"}

    #Pull profile values from the spreadsheet

    #Get profile percentages (tier 1 is 50% of total, tier 2 is 25%, tier 3 is 25%)
    [int]$NumTier1MBX = ($InputSheet.Range("NumTier1MBX")).Text
    [int]$NumTier2MBX = ($InputSheet.Range("NumTier2MBX")).Text
    [int]$NumTier3MBX = ($InputSheet.Range("NumTier3MBX")).Text
    [int]$NumTier4MBX = ($InputSheet.Range("NumTier4MBX")).Text

    $TotalMBX = ($NumTier1MBX + $NumTier2MBX + $NumTier3MBX + $NumTier4MBX)
    Write-VerboseOutput("Number of mailboxes in all tiers: " + $TotalMBX)
    $script:P1Percent = [math]::Round(([float]($NumTier1MBX/$TotalMBX) * 100), 2)
    Write-VerboseOutput("Percentage of mailboxes in tier 1: " + $script:P1Percent)
    $script:P2Percent = [math]::Round(([float]($NumTier2MBX/$TotalMBX) * 100), 2)
    Write-VerboseOutput("Percentage of mailboxes in tier 2: " + $script:P2Percent)
    $script:P3Percent = [math]::Round(([float]($NumTier3MBX/$TotalMBX) * 100), 2)
    Write-VerboseOutput("Percentage of mailboxes in tier 3: " + $script:P3Percent)
    $script:P4Percent = [math]::Round(([float]($NumTier4MBX/$TotalMBX) * 100), 2)
    Write-VerboseOutput("Percentage of mailboxes in tier 4: " + $script:P4Percent)

    #Get Messages/sent received per day by profile.  It's stored in format "50 messages" so we have to extract the number only.
    $script:P1Messages = [int]((($InputSheet.Range("Tier1SendReceive")).Text).Split(" "))[0]
    Write-VerboseOutput("Profile 1 Send Receive Messages Per day: " + $script:P1Messages)
    $script:P2Messages = [int]((($InputSheet.Range("Tier2SendReceive")).Text).Split(" "))[0]
    Write-VerboseOutput("Profile 2 Send Receive Messages Per day: " + $script:P2Messages)
    $script:P3Messages = [int]((($InputSheet.Range("Tier3SendReceive")).Text).Split(" "))[0]
    Write-VerboseOutput("Profile 3 Send Receive Messages Per day: " + $script:P3Messages)
    $script:P4Messages = [int]((($InputSheet.Range("Tier4SendReceive")).Text).Split(" "))[0]
    Write-VerboseOutput("Profile 4 Send Receive Messages Per day: " + $script:P4Messages)

    #Get mailbox multiplier
    $script:P1Multiplier = [float]($InputSheet.Range("Tier1MCFactor")).Text
    Write-VerboseOutput("Profile 1 multiplication factor: " + $script:P1Multiplier)
    $script:P2Multiplier = [float]($InputSheet.Range("Tier2MCFactor")).Text
    Write-VerboseOutput("Profile 2 multiplication factor: " + $script:P2Multiplier)
    $script:P3Multiplier = [float]($InputSheet.Range("Tier3MCFactor")).Text
    Write-VerboseOutput("Profile 3 multiplication factor: " + $script:P3Multiplier)
    $script:P4Multiplier = [float]($InputSheet.Range("Tier4MCFactor")).Text
    Write-VerboseOutput("Profile 4 multiplication factor: " + $script:P4Multiplier)

    #Get percent of users the multiplier applies to
    $script:P1MultiplierPercent = [int]($InputSheet.Range("Tier1UserMFPercentage")).Text.TrimEnd('%')
    Write-VerboseOutput("Profile 1 multiplication factor user percentage: " + $script:P1MultiplierPercent)
    $script:P2MultiplierPercent = [int]($InputSheet.Range("Tier2UserMFPercentage")).Text.TrimEnd('%')
    Write-VerboseOutput("Profile 2 multiplication factor user percentage: " + $script:P2MultiplierPercent)
    $script:P3MultiplierPercent = [int]($InputSheet.Range("Tier3UserMFPercentage")).Text.TrimEnd('%')
    Write-VerboseOutput("Profile 3 multiplication factor user percentage: " + $script:P3MultiplierPercent)
    $script:P4MultiplierPercent = [int]($InputSheet.Range("Tier4UserMFPercentage")).Text.TrimEnd('%')
    Write-VerboseOutput("Profile 4 multiplication factor user percentage: " + $script:P4MultiplierPercent)

    #Get SPECInt Rating.  We assume this is the primary datacenter.  If the -SecondaryDataCenter switch was passed pull the SPEC rating for the secondary data center instead.
    if($SecondaryDataCenter)
    {
        $TotalSPECRating = [int]($InputSheet.Range("numSpecRateSDC")).Text
        $script:PAllSpecRating = $TotalSPECRating/$script:NumberOfCores
    }
    else
    {
        $TotalSPECRating = [int]($InputSheet.Range("numSpecRatePDC")).Text
        $script:PAllSpecRating = $TotalSPECRating/$script:NumberOfCores
    }
    Write-VerboseOutput("Total SPEC Rating: " + $TotalSPECRating)
    Write-VerboseOutput("Per-Core SPEC Rating: " + $script:PAllSpecRating)
}

function Set-ActiveMegacyclesBasedOnProfile
{
    Write-VerboseOutput("Calling Set-ActiveMegacyclesBasedOnProfile")

    #Convert messages/day number or megacycle requirement
    $script:P1ActiveMegacycles = Get-ActiveMailboxRequiredMegacycles($script:P1Messages)
    Write-VerboseOutput("Active Megacycles for Profile 1: " + $script:P1ActiveMegacycles)
    $script:P1PassiveMegacycles = Get-PassiveMailboxRequiredMegacycles($script:P1Messages)
    Write-VerboseOutput("Passive Megacycles for Profile 1: " + $script:P1PassiveMegacycles)

    #only bother with profiles 2-4 if the percent of mailboxes is greater than 0
    if($script:P2Percent -gt 0)
    {
        $script:P2ActiveMegacycles = Get-ActiveMailboxRequiredMegacycles($script:P2Messages)
        Write-VerboseOutput("Active Megacycles for Profile 2: " + $script:P2ActiveMegacycles)
        $script:P2PassiveMegacycles = Get-PassiveMailboxRequiredMegacycles($script:P2Messages)
        Write-VerboseOutput("Passive Megacycles for Profile 2: " + $script:P2PassiveMegacycles)
    }
    if($script:P3Percent -ne 0)
    {
        $script:P3ActiveMegacycles = Get-ActiveMailboxRequiredMegacycles($script:P3Messages)
        Write-VerboseOutput("Active Megacycles for Profile 3: " + $script:P3ActiveMegacycles)
        $script:P3PassiveMegacycles = Get-PassiveMailboxRequiredMegacycles($script:P3Messages)
        Write-VerboseOutput("Passive Megacycles for Profile 3: " + $script:P3PassiveMegacycles)
    }
    if($script:P3Percent -ne 0)
    {
        $script:P4ActiveMegacycles = Get-ActiveMailboxRequiredMegacycles($script:P4Messages)
        Write-VerboseOutput("Active Megacycles for Profile 4: " + $script:P4ActiveMegacycles)
        $script:P4PassiveMegacycles = Get-PassiveMailboxRequiredMegacycles($script:P4Messages)
        Write-VerboseOutput("Passive Megacycles for Profile 4: " + $script:P4PassiveMegacycles)
    }
}

#On multi proc boxes, WMI reports number of cores and megacycles per core as an array value for each proc such as @(8,8,8,8) instead of 32.
#It can also put the results of Get-WmiObject Win32_Processor into an array of Win32_Processor objects depending on the hardware setup.  
#Need to normalize these numbers to avoid errors.
function Normalize-ProcessorInfo
{
    Write-VerboseOutput("Calling Normalize-ProcessorInfo")

    #Handle single and multi proc machines slightly differently due to the way Win32_Processor returns the data
    if($system.NumberOfProcessors -gt 1)
    {
        #Total cores in all processors
        foreach($processor in $proc)
        {
            $coresum += $processor.NumberOfCores
            $logicalsum += $processor.NumberOfLogicalProcessors
            if($processor.CurrentClockSpeed -lt $processor.MaxClockSpeed)
            {
                $script:CurrentMegacycles = $processor.CurrentClockSpeed
                $script:ProcessorIsThrottled = $true
            }
        }
        $script:NumberOfCores = $coresum
        $script:NumberOfLogicalProcessors = $logicalsum
     
        #all processors should be the same speed and type so take the description and Max Speed of the first processor
        $script:ProcessorName = $proc[0].Name
        $script:MegacyclesPerCore = $proc[0].MaxClockSpeed
    }
    else #single processor machine
    {
        $script:NumberOfCores = $proc.NumberOfCores
        $script:NumberOfLogicalProcessors = $proc.NumberOfLogicalProcessors
        $script:MegacyclesPerCore = $proc.MaxClockSpeed
        $script:ProcessorName = $proc.Name
        if($proc.CurrentClockSpeed -lt $proc.MaxClockSpeed)
        {
            $script:CurrentMegacycles = $proc.CurrentClockSpeed
            $script:ProcessorIsThrottled = $true
        }
    }

    #We need processor count, cores, logical processors, and megacycles to continue.  If one of these is missing, exit the script.
    if(($script:NumberOfCores -eq $null) -or ($script:NumberOfLogicalProcessors -eq $null) -or ($script:MegacyclesPerCore -eq $null))
    {
        Write-Red("Processor information could not be read.  Exiting script.")
        Exit
    }
}

#Main script execution
Write-VerboseOutput("Calling Main Script Execution")

if(-not (Test-Path $OutputFilePath))
{
    Write-Host "Invalid value specified for -OutputFilePath." -ForegroundColor Red
    Exit
}

#Normalize processor values
Normalize-ProcessorInfo

#Populate server role
$ServerRole = Get-ServerRole

#Display system and processor information to console
Write-SystemInformationToConsole

#Load values from the sizing calculator results
if($CalculatorFile -ne $null)
{
    Set-ProfileValuesFromSpreadsheet
}

#Convert messages/day to megacycles needed
Set-ActiveMegacyclesBasedOnProfile

#Validate profile percentage numbers
if([math]::Round(($script:P1Percent + $script:P2Percent + $script:P3Percent + $script:P4Percent), 0) -ne 100)
{
    Write-Red("Total of the profile percentage values must equal 100")
    exit
}

#Calculate needed megacycles for current server configuration
Check-NeededMegacycles

#Finish
Write-Grey("Output file written to " + $OutputFullPath)












