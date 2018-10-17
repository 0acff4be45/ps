
[CmdletBinding()]

Param(
	[ValidateScript({$_ -like "*domain.ac.uk"})] 
   	[Parameter(Mandatory=$False,HelpMessage="Server FQDN")]
   	[string]$Server = "$($env:computername).domain.ac.uk",
	
   	[ValidateScript({$_ -like "*domain.ac.uk"})] 
	[Parameter(Mandatory=$True,HelpMessage="Target FQDN for active DBs")]
   	[string]$Target,
	
	[Parameter(Mandatory=$False,HelpMessage="LogFile Path")]
   	[string]$LogPath = "D:\Scripts\",
	
	[Parameter(Mandatory=$False,HelpMessage="LogFile Path")]
   	[string]$LogFile = "StartExchangeMaintenance.log"
	
)

$TranscriptPath = $LogPath + "Transcript - ExchangeStartMaintenance.log"
Start-Transcript -path $TranscriptPath -Force

#Setup variables
$LogFullPath = $LogPath + $LogFile
$MessagesRedirected = "OK"
$QueuesEmpty = "OK"
$SuspendDAG = "OK"
$ServicesStopped = "OK"
$ServicesRestarted = "OK"
$ServerOffline = "OK"

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

Function Get-AllQueues {
	$AllQueues = Get-Queue -Server $Server | ?{$_.Identity -notlike "*\Poison" –and $_.Identity -notlike"*\Shadow\*"} | select identity, messagecount 
	$Total = 0
	ForEach ($Item in $AllQueues)
		{$Total += $Item.MessageCount}
	$Total
}
cls

# Connect to Exchange remotely - not used in this case
#$ExSess = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri 'http://$Server/powershell/?serializationlevel=full' 
#Import-PSSession $ExSess -AllowClobber| Out-Null

# Load the Exchange snapin locally (if it's not already present)
LoadExchangeSnapin

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
WriteLog "Server is $Server"
WriteLog "Target for mail queues is $Target"
WriteLog " "
		
### Start Maintenance Procedures ###
WriteLog "--- Starting maintenance ---"

# Drain Hub Transport queues
WriteLog "Putting queues in maintenance mode..."
Set-ServerComponentState $Server –Component HubTransport –State Draining –Requester Maintenance -confirm:$false 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error setting server component state" 
	  $MessagesRedirected = "Failed setting component state"
	}
else
	{WriteLog "Set server component state to Draining on $server"}

Redirect-Message -Server $Server -Target $Target -confirm:$false 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error redirecting messages from $Server to $Target" 
	  $MessagesRedirected = "Failed redirecting messages to $Target"
	}
else
	{WriteLog "Redirected messages from $Server to $Target" }

	# Wait for queues to clear
WriteLog " "
WriteLog "Waiting for queues to drain. Press CTRL-C to abort..."
$Time = 0

try
{
# Catch CTRL-C
[console]::TreatControlCAsInput = $true

    While ($(Get-AllQueues) -ne 0) 
	    {

		    Start-sleep -Seconds 1

            # Only process main code every second
            $Time++
            WriteLog "$(Get-AllQueues) items remaining: Waited $Time sec (time is $(get-date))."
		    
            # If Get-AllQueues isn't 0 after 2 minutes, give up.
            If ($Time -ge 120) 
			    { WriteLog "Queues not clear after 2 minutes. Giving up waiting..."
			        break;
			    }

            # Check for CTRL-C and handle it
                if ([console]::KeyAvailable) {
                    $key = [system.console]::readkey($true)
                    if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                        $breaking = $true
                        Write-Host ("CTRL-C pressed. Wait cancelled after $Time seconds (time is $(get-date))")
                        break
                    }
                }
	    }
}
finally
{
    # Clean up after try block
    [console]::TreatControlCAsInput = $false
}

If ($(Get-AllQueues) -eq 0)
	{ WriteLog "Queues drained. Moving on..." }
else 
	{WriteLog "$(Get-AllQueues) items remaining in the mail queues when wait abandoned."
	 $QueuesEmpty = "$(Get-AllQueues) items remaining in the mail queues when wait abandoned."
	}

WriteLog " "
WriteLog "Moving active databases to another server and suspending DAG membership"
WriteLog " "
try {
	C:\Program` Files\Microsoft\Exchange` Server\V15\Scripts\StartDagServerMaintenance.ps1 -server $server –overrideMinimumTwoCopies $true
}
Catch {
		WriteLog "ERROR: Error suspending DAG and moving DBs" 
		$SuspendDAG = "Script failed"
	  }
If (!$?) 
	{ WriteLog "ERROR: Error suspending DAG and moving DBs" 
	  $SuspendDAG = "Script failed"
	}
else
	{WriteLog "DAG server maintenance script complete" }

# Restarting services
WriteLog " "
WriteLog "Restarting MSExchangeTransport service on $Server"
Restart-Service MSExchangeTransport 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error restarting service" 
	  $ServicesRestarted = "Failed" }
else
	{WriteLog "Service restarted"}

WriteLog " "
WriteLog "Restarting MSExchangeFrontEndTransport service on $Server"
Restart-Service MSExchangeFrontEndTransport 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error restarting service" 
	  $ServicesRestarted = "Failed" }
else
	{WriteLog "Service restarted"}

# Set server wide offline
WriteLog " "
WriteLog "Setting Server Wide Offline on $Server"
Set-ServerComponentState $env:computername –Component ServerWideOffline –State InActive –Requester Maintenance -confirm:$false 
If (!$?) 
	{ WriteLog "ERROR: Error setting server wide offline"  
	  $ServerOffline = "Failed" }
else
	{WriteLog "Server Wide Offline set"}
	
#Stop Check_MK agent
WriteLog " "
WriteLog "Stopping check_MK agent on $Server"
Stop-Service check_mk_agent 
# Error Check
If (!$?) 
	{ WriteLog "ERROR: Error stopping service" 
	  $ServicesStopped = "Failed" }
else
	{WriteLog "Service stopped"}

# Clean up
$CurTime = Get-Date 
cls
WriteLog " "
WriteLog "Script finished at $CurTime"
WriteLog " "
WriteLog "--- Activity report ---"
WriteLog "Messages Redirected: $MessagesRedirected"
WriteLog "Queues Empty: $QueuesEmpty"
WriteLog "Move active DBs and suspend DAG: $SuspendDAG" 
WriteLog "Transport services restarted: $ServicesRestarted"
WriteLog "CheckMK service stopped: $ServicesStopped"
WriteLog "Server Wide Offline: $ServerOffline"
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