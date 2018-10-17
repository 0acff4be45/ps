

[CmdletBinding()]

Param(
	[ValidateScript({$_ -like "*domain.com"})] 
   	[Parameter(Mandatory=$False,HelpMessage="Server FQDN")]
   	[string]$Server = "$($env:computername).domain.com",
	
	[Parameter(Mandatory=$False,HelpMessage="LogFile Path")]
   	[string]$LogPath = "D:\Scripts\",
	
	[Parameter(Mandatory=$False,HelpMessage="DAG Name")]
   	[string]$DAGName = "DAGEX13",
	
	[Parameter(Mandatory=$False,HelpMessage="LogFile Path")]
   	[string]$LogFile = "StopExchangeMaintenance.log"
	
)

$TranscriptPath = $LogPath + "Transcript - ExchangeStopMaintenance.log"
Start-Transcript -path $TranscriptPath -Force

#Setup variables
$LogFullPath = $LogPath + $LogFile
$ResumeDAG = "OK"
$DBActivationFlag = "OK"
$HAComponentState = "OK"
$MessagesResumed = "OK"
$ServicesStarted = "OK"
$ServicesRestarted = "OK"
$ServerOnline = "OK"
$RebalanceDBs = "OK"

function WriteLog {
	param(
	$LogData
	)
	$LogData | Out-File -FilePath $LogFullPath -Append
	Write-Host $LogData 
}

function LoadExchangeSnapin
{
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:Stop
    }
}

Function Show-ExForest { Set-ADServerSettings -ViewEntireForest $True }

cls

# Load the Exchange snapin locally (if it's not already present)
LoadExchangeSnapin

# Load MS DAG functions library
. "c:\Program Files\Microsoft\Exchange Server\V15\Scripts\DagCommonLibrary.ps1"
Test-RsatClusteringInstalled

# Allow an FQDN to be passed in, but strip it to the short name.
$shortServerName = $server;
if ( $shortServerName.Contains( "." ) )
{
	$shortServerName = $shortServerName -replace "\..*$"
}

try {Show-ExForest} catch { write-host "Warning: ViewEntireForest failed."}

# Store script location in a variable.
$ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition

# Clear PowerShell error buffer
$error.clear()

# Initialise logfile
$CurTime = Get-Date 
Write-Host "Log file is $LogFullPath"
Write-Host " "
try {
        "Script started at $CurTime" | Out-File -FilePath $LogFullPath }
catch {
        #cls
		Write-Error "`nERROR: Log file creation failed with this error: `n`n"
		exit }
		
WriteLog " "
WriteLog "--- Script parameters ---"
WriteLog "Server is $Server, DAG is $DAGName"
WriteLog " "
		
### Start Maintenance Procedures ###
WriteLog "--- Starting maintenance ---"

#Start Check_MK agent
WriteLog " "
WriteLog "Starting check_MK agent on $Server"
Start-Service check_mk_agent 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error starting service" 
	  $ServicesStarted = "Check_MK_Agent failed to start" }
else
	{WriteLog "Check_MK_Agent service started"}

# Set server wide offline to active
WriteLog " "
WriteLog "Setting Server Wide Offline to active on $Server"
Set-ServerComponentState $shortServerName –Component ServerWideOffline –State Active –Requester Maintenance -confirm:$false 
If (!$?) 
	{ WriteLog "ERROR: Error setting server wide offline state"  
	  $ServerOnline = "Failed" }
else
	{WriteLog "Server Wide Offline state set to active"}
	
# Take DAG out of maintenance mode
WriteLog " "
WriteLog "Resuming DAG membership for $shortServerName in cluster $DAGName"
WriteLog " "

$outputStruct = Call-ClusterExe -dagName $dagName -serverName $shortServerName -clusterCommand "node $shortServerName /resume"
$LastExitCode = $outputStruct[ 0 ];

# 0 is success, 5058 is ERROR_CLUSTER_NODE_NOT_PAUSED, 1753 is EPT_S_NOT_REGISTERED, 1722 is RPC_S_SERVER_UNAVAILABLE
if ( $LastExitCode -ne 0 )
{
	WriteLog ("Cluster resume failed: cluster.exe returned error $LastExitCode")
    $ResumeDAG = "Failed with cluster.exe error $LastExitCode"
}
else
{
	WriteLog ("Cluster $DAGName was resumed successfully on $Server")
}

# Unsetting DatabaseCopyActivationDisabledAndMoveNow flag
WriteLog " "
WriteLog "Unsetting DatabaseCopyActivationDisabledAndMoveNow flag on $Server..."
Set-MailboxServer -Identity $shortServerName -DatabaseCopyActivationDisabledAndMoveNow:$false
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error setting DatabaseCopyActivationDisabledAndMoveNow flag" 
	  $DBActivationFlag = "Failed setting DatabaseCopyActivationDisabledAndMoveNow flag"
	}
else
	{WriteLog "DatabaseCopyActivationDisabledAndMoveNow flag set to false on $server"}

# Setting DB Auto Activation Policy to unrestricted
WriteLog " "
WriteLog "Setting DatabaseCopyAutoActivationPolicy to Unrestricted on $Server..."
Set-MailboxServer -Identity $shortServerName -DatabaseCopyAutoActivationPolicy:Unrestricted
If (!$?) 
	{ WriteLog "ERROR: Error setting DatabaseCopyAutoActivationPolicy" 
	  $DBActivationFlag += ", Failed setting DatabaseCopyAutoActivationPolicy"
	}
else
	{WriteLog "Set DatabaseCopyAutoActivationPolicy to unrestricted on $server"}

# Set HA component state to active
WriteLog " "
WriteLog "Setting HighAvailability component state to active..."
Set-ServerComponentState $server -Component "HighAvailability" -Requester "Maintenance" -State Active
If (!$?) 
	{ WriteLog "ERROR: Error setting component state" 
	  $HAComponentState = "Failed setting HighAvailability component state - try manually"
	}
else
	{WriteLog "HighAvailability component state set to Active on $server"}

# Best effort resume copies for backward compatibility of machines who were activation suspended with the previous build.
WriteLog " "
WriteLog "Best effort attempt to resume DB copies that were activation suspended with a previous build..."
try
{
	$databases = Get-MailboxDatabase -Server $shortServerName | where { $_.ReplicationType -eq 'Remote' }

	if ( $databases )
	{
		# 1. Resume database copy. This clears the ActivationOnly suspension.
		foreach ( $database in $databases )
		{
			WriteLog ("Resuming mailbox database copying on $($database.Name)\\$shortServerName. This clears the Activation Suspended state")
			Resume-MailboxDatabaseCopy "$($database.Name)\$shortServerName" -Confirm:$false
		}
	}
}
catch
{
	WriteLog ("Best effort attempt at resuming database copies on server $shortServerName failed with error $($_)")
}
WriteLog "Finished best effort resume"

# Reactivate Hub Transport queues
WriteLog " "
WriteLog "Taking queues out of maintenance mode..."
Set-ServerComponentState $shortServerName –Component HubTransport –State Active –Requester Maintenance -confirm:$false 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error setting hub transport component state" 
	  $MessagesResumed = "Failed setting hub transport component state"
	}
else
	{WriteLog "Set hub transport component state to active on $server"}
	
# Restarting services
WriteLog " "
WriteLog "Restarting MSExchangeTransport service on $Server"
Restart-Service MSExchangeTransport 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error restarting service" 
	  $ServicesRestarted = " MSExchangeTransport Failed" }
else
	{WriteLog "Service restarted"}

WriteLog " "
WriteLog "Restarting MSExchangeFrontEndTransport service on $Server"
Restart-Service MSExchangeFrontEndTransport 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error restarting service" 
	  $ServicesRestarted += ", MSExchangeFrontEndTransport Failed" }
else
	{WriteLog "Service restarted"}
	
# Rebalance active DBs accross DAG
WriteLog " "
WriteLog "Rebalancing DBs: attempt 1"
WriteLog " "
try {
	C:\Program` Files\Microsoft\Exchange` Server\V15\Scripts\RedistributeActiveDatabases.ps1 -BalanceDbsByActivationPreference -confirm:$false 
}
Catch {
		WriteLog "ERROR: Error rebalancing DBs" 
		$RebalanceDBs = "Script failed"
	  }
If (!$?) 
	{ WriteLog "ERROR: Error rebalancing DBs" 
	  $RebalanceDBs = "Script failed"
	}
else
	{WriteLog "DBs rebalanced: attempt 1" }

WriteLog "Waiting 60 seconds before trying to rebalance DAGs again..."

# Wait 1 minute and rebalance active DBs accross DAG again (MS' script doesn't always work first time)
$WaitDelay = 60
While ($WaitDelay -gt 0)
	{
		Write-Progress -activity "Waiting 60 seconds to rebalance DAGs again" -status "Time remaining: " -PercentComplete $((100/60)*$WaitDelay) -CurrentOperation ("$WaitDelay seconds remaining...")
		Start-sleep -s 1
		$WaitDelay --
	}
# Clear progress bar
Write-Progress -activity "Done" -Status "Done" -Completed

WriteLog " "
WriteLog "Rebalancing DBs: attempt 2"
WriteLog " "
try {
	C:\Program` Files\Microsoft\Exchange` Server\V15\Scripts\RedistributeActiveDatabases.ps1 -BalanceDbsByActivationPreference -confirm:$false 
}
Catch {
		WriteLog "ERROR: Error rebalancing DBs" 
		$RebalanceDBs = "Script failed on run 2"
	  }
If (!$?) 
	{ WriteLog "ERROR: Error rebalancing DBs" 
	  $RebalanceDBs = "Script failed on run 2"
	}
else
	{WriteLog "DBs rebalanced: attempt 2" }

# Clean up
$CurTime = Get-Date 
cls
WriteLog " "
WriteLog "Script finished at $CurTime"
WriteLog " "
WriteLog "--- Activity report ---"
WriteLog "CheckMK service started: $ServicesStarted"
WriteLog "Server Wide Offline set to Active: $ServerOnline"
WriteLog "Resumed suspended server in DAG: $ResumeDAG" 
WriteLog "HighAvailability set to Active: $HAComponentState"
WriteLog "DB replication activated: $DBActivationFlag"
WriteLog "Message queues resumed: $MessagesResumed"
WriteLog "Transport services restarted: $ServicesRestarted"
WriteLog "Rebalanced Active DBs in DAG: $RebalanceDBs"
WriteLog "-----------------------"
WriteLog " "

if ($error) 
	{
		WriteLog "!!! $($error.count) errors were encountered. Please check log for details."
		"Erorrs are shown below:" | Out-File -FilePath $LogFullPath -Append
		$error | Out-File -FilePath $LogFullPath -Append
	} 
else 
	{WriteLog "No errors encountered"}



Stop-Transcript