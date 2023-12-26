#################################################
# Script Creator:
# Paul Arquette
#
# Script Purpose:
# Force SCCM Checkin For Specified Boxes
#
# Date Created:
# September 26, 2023
#
# Date Last Modified:
# December 26, 2023  (Modified For Public Github)
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
$computers = @()
$EmailFrom = "Server@domain.com"
$EmailTo = "recipient@domain.com"
$EmailSubject = "Windows Servers (TIME) MECM Force Install Patches"
$EmailBody = "The script to force Windows Prod Servers at 3AM to Install Available Patches Completed."
$SMTPServer = "smtp.domain.com"
$SMTPPort = "25"
$ComputerADGroup = "MECM_PatchMgmt_Servers_3AM"

#################################################
# Start Script Programming Here
#################################################

$computers = Get-ADGroupMember -Identity $ComputerADGroup

ForEach ($server in $computers)
{
    $servername = $server.name

    Write-Host "Working on server $servername"
    $s = New-PSSession -computerName $servername
    Invoke-Command -Session $s -Scriptblock {$UpdatesAvailable = Get-CimInstance -Namespace 'root\ccm\ClientSDK' -ClassName 'CCM_SoftwareUpdate' -Filter 'EvaluationState = 0 OR EvaluationState = 1 OR EvaluationState = 13'}
    Invoke-Command -Session $s -Scriptblock {Invoke-CimMethod -Namespace 'root\ccm\ClientSDK' -ClassName 'CCM_SoftwareUpdatesManager' -MethodName 'InstallUpdates' -Arguments @{CCMUpdates = [ciminstance[]]$UpdatesAvailable}}

    Remove-PSSession $s  
}

# Send Out E-mail if Necessary
#################################################
$From = $EmailFrom
$To = $EmailTo
$Subject = $EmailSubject
$Body = $EmailBody
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort
