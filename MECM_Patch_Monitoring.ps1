#################################################
# Script Creator:
# Paul Arquette
#
# Script Purpose:
# Check Windows Servers During Patching, 
# Report Last Boot Time, Force SCCM Check When Required, 
# and Restart Windows Services if Required 
#
# Date Created:
# September 13, 2023
#
# Date Last Modified:
# December 26, 2023 (Modified for Public Github)
#################################################


## HTML STYLING
####################################################################################
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
#####################################################################################

# Set Global Variables & Input Files
#################################################
$online = @()
$offline = @()
$reportboot = @()
$reportsvc = @()

#Email Variables
$From = "server@domain.com"
$To = "recipient@domain.com"
$Subject = "Server Patch Report"
$SMTPServer = "smtp.domain.com"
$SMTPPort = "25"

$ADComputerGroupOne = "MECM_PatchMgmt_Test_9PM"
$ADComputerGroupTwo = "MECM_PatchMgmt_Test_10PM"
$ADComputerGroupThree = "MECM_PatchMgmt_Test_11PM"

# Functions - Declare Functions
#################################################
function WriteLog
{
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
}


#################################################
# Start Script Programming Here
#################################################

$computers = Get-ADGroupMember -Identity $ADComputerGroupOne
$computers += Get-ADGroupMember -Identity $ADComputerGroupTwo
$computers += Get-ADGroupMember -Identity $ADComputerGroupThree


#Check For Online/Offline
###########################
ForEach ($server in $computers)
{
    $servername = $server.name
    $netconn_status = Test-NetConnection -ComputerName $servername -Port 5985

    if ($netconn_status.TcpTestSucceeded -eq "True")
    {
        $online += $servername
    } else {
        $offline += $servername           
    }
}


#Check For Last Server Boot Time
########################

ForEach ($server in $online)
{ 
    $curdateboot = Get-Date
    $servername = $server
    $bootuptime = Get-CimInstance -ClassName win32_operatingsystem -ComputerName $servername | select csname, lastbootuptime
    $lastboot = $bootuptime.lastbootuptime

    if ( ($curdateboot -gt $lastboot.AddMinutes(25) -and ($lastboot -gt $curdateboot.AddHours(-4)) )) #If server rebooted at least 25 minutes ago but not more than 4 hours ago, force check in
    {
        #Force SCCM Check-In After Reboot
        Invoke-Command -computername $servername -scriptblock {Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"}
        Invoke-Command -computername $servername -scriptblock {Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000108}"}
        Invoke-Command -computername $servername -scriptblock {Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}"}
        $reportboot += New-Object PSObject -Property @{Name=$servername;LastBoot=$lastboot;SCCMForceCheckIn="Yes"}
    } else {
        $reportboot += New-Object PSObject -Property @{Name=$servername;LastBoot=$lastboot;SCCMForceCheckIn="No"}
    }       
        
}

ForEach ($server in $offline)
{
    $reportboot += New-Object PSObject -Property @{Name=$server;LastBoot="_UNAVAILABLE"}
}


#Check For Services
########################

$jobC = Invoke-Command -computername $online -scriptblock {Get-Service | where{$_.StartType -eq "Automatic"} | Select Name, Starttype, status} -AsJob
            
Start-Sleep -Seconds 60

$jobs = $jobC.childjobs

    ForEach ($job in $jobs)
    {
        if ($job.state -eq "Completed")
        {
            $jobid = $job.id
            $badsvcs = Receive-Job -ID $jobid |Where {$_.Status -notlike "Running"}

            ForEach ($badsvc in $badsvcs)
            {
                $badsvcname = $badsvc.Name

                if (($badsvcname -eq "sppsvc") -or ($badsvcname -eq "WbioSrvc") -or ($badsvcname -eq "RemoteRegistry") -or ($badsvcname -eq "CDPSvc") -or ($badsvcname -eq "MSExchangeNotificationsBroker") -or ($badsvcname -eq "edgeupdate"))
                {
                    #Do Nothing With These Services (Whitelisted)
                } else {
                    $servernameSVC = $badsvc.PSComputerName
                    $bootuptime = Get-CimInstance -ClassName win32_operatingsystem -ComputerName $servernameSVC | select csname, lastbootuptime
                    $lastboot = $bootuptime.lastbootuptime
                    $currDT = Get-Date

                    if ( ($currDT -gt $lastboot.AddMinutes(25) -and ($lastboot -gt $currDT.AddHours(-4)) )) #If system booted between 25 minutes ago and 4 hours ago, force restart services.
                    {                        
                        #Restart Services
                        Invoke-Command -computername $servernameSVC -scriptblock {Start-Service -Name $Using:badsvcname}
                        $reportsvc += New-Object PSObject -Property @{Name=$badsvc.PSComputerName;ServicesNotStarted=$badsvc.Name;ServiceRestartAttempted="Yes"}
                    }else {
                        $reportsvc += New-Object PSObject -Property @{Name=$badsvc.PSComputerName;ServicesNotStarted=$badsvc.Name;ServiceRestartAttempted="No"}
                    }
                }
            }

        } else {
            $jobid = $job.id
            $location = $job.location
            Write-Host "The following job did not complete: $jobid"
            $reportsvc += New-Object PSObject -Property @{Name=$location;ServicesNotStarted="Job Not Completed"}
        }
    }              

ForEach ($server in $offline)
{
    $reportsvc += New-Object PSObject -Property @{Name=$server;ServicesNotStarted="_UNAVAILABLE"}
}



# Send Out E-mail if Necessary
#################################################
$emailResponse = $reportboot |Select-Object Name,LastBoot,SCCMForceCheckIn |Sort-Object LastBoot |ConvertTo-Html -Head $style
$emailResponse2 = $reportsvc |Select-Object Name,ServicesNotStarted,ServiceRestartAttempted |Sort-Object Name |ConvertTo-Html -Head $style

$Body = @"
<b>Last Boot Report</b>
$emailResponse
<br>
<b>Services Not Started Report</b>
$emailResponse2
"@
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort
