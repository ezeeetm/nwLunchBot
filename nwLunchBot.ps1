<#
.SYNOPSIS
    Scrapes Nationwide's cafeteria menus for restaurants, and sends emails/SNS to subscribers.  Run as a Windows scheduled task an hour before lunch.
    													
.NOTES
	DEPENDS: None
    COMMENTS: 
	    - for SNS email domains:
            Alltel [insert 10-digit number] @message.alltel.com
            AT&T [insert 10-digit number] @txt.att.net
            Boost Mobile [insert 10-digit number] @myboostmobile.com
            Sprint [insert 10-digit number] @messaging.sprintpcs.com
            T-Mobile [insert 10-digit number] @tmomail.net
            US Cellular [insert 10-digit number] @email.uscc.net
            Verizon [insert 10-digit number] @vtext.com
            Virgin Mobile [insert 10-digit number] @vmobl.com
            Republic Wireless [insert 10-digit number] @text.republicwireless.com 
            Google Fi [insert 10-digit number] @msg.fi.google.com
#>


# config
$subscribers=@(
'1234567899@txt.att.net',
'somegmail@gmail.com'
)

$locations = @(
@('3747','Plaza One'),
@('3743','Plaza Three'),
@('3742','Rings Rd')
)

$restaurants = @(
'Panda Express',
'Tavolino Pasta Bar',
'Short North Bagel',
'Hot Heads',
'Broad Street Philly''s',
'Claddagh',
'Zoca',
'Bob Evans',
'Olive Tree Cafe',
'Fusian',
'Smokey Bones',
'Curry N Grill',
'Taste of Sicily',
'Skyline Chili',
'Schmidt''s',
'Cumin',
'Lemongrass',
'ZOCA!',
'Mazah',
'J Gumbo',
'Olive Tree',
'Greek Street'
)

# Be sure to enable this for gmail accounts: https://www.google.com/settings/security/lesssecureapps.  System.Net.Mail.SmtpClient sucks.
# This only needs to be enabled for the SENDING gmail account that is specfied by the u/p below.  The subscribers do *not* need to enable this.
$SMTPServer = "smtp.gmail.com"
$SMTPPort = "587"
$Username = "sendingGmailAccount@gmail.com"
$Password = 'sendingGmailAccountPassWord'


# logging & error handling
clear
# create .\<ScriptName>Logs\ dir if it does not exit
$path = Split-Path $MyInvocation.MyCommand.Path
Set-Location -Path $path
$ScriptName = "nwLunchBot"
$logPath = ".\"+$ScriptName+"Logs"
if((Test-Path $logPath) -eq $false){md -Name $logPath | out-null} #suppress stdout

# start new log file w/ timestamp filename and date header
$logStartDateTime = Get-Date -f yyyy-MM-dd
$logHeader = "Log File Started: "+(Get-Date).ToString()+"`r`n-------------------------------------"
$logHeader | Out-File ".\$logPath\$logStartDateTime.log" -Append

function log ($logLevel, $logData)
{
    $logEntry = (Get-Date).ToString()+"`t"+$logLevel+"`t"+$logData
    # Write-Host $logEntry
    $logEntry | Out-File ".\$logPath\$logStartDateTime.log" -Append
}

function HandleException($function,$exceptionMessage)
{
    log WARNING "Exception in $($function): $exceptionMessage"
}
# /logging & error handling


function GetMsg ($locations)
{
    try
    {
        log INFO "  getMsg: STARTED"
        $msg = ""
        $today = Get-Date -Format "dddd,MMMM d,yyyy"
        #$today = (get-date).AddDays(2).ToString("dddd,MMMM d,yyyy") #use this line instead when testing/debugging on days when there is no menu
        $found = $false

        foreach ($location in $locations)
        {
            $locationId = $location[0]
            $locationName = $location[1]

            $resp = Invoke-WebRequest "http://www.aramarkcafe.com/layouts/canary_2010/locationhome.aspx?locationid=$locationId"
            $todaysMenuUrl = "http://www.aramarkcafe.com/layouts/canary_2010/$(($resp.Links | where {$_.innerHTML -eq $today}).href)".Replace("&amp;","&")
            
            log INFO "    $locationName-$locationId todaysMenuUrl: $todaysMenuUrl"

            $resp = (Invoke-WebRequest $todaysMenuUrl).content

            $msg += "$($locationName):`n"
            foreach ($restaurant in $restaurants)
            {
                if ($resp.contains($restaurant))
                {
                    $msg += "  $restaurant`n"
                    $found = $true
                    log INFO "      found $restaurant"
                }
            }

            $msg += "`n"
        }

        if (!$found)
        {
            $msg = $null
        }
    }
    Catch [Exception]
    {
        $exMsg = $_.Exception.Message
        HandleException "getMsg" $exMsg
    }

    log INFO "  getMsg: COMPLETED"
    return $msg
}


function sendMsg ($subscribers,$msg)
{
    $subDate = Get-DAte -Format "MMMM dd"
    foreach ($subscriber in $subscribers)
    {
        Try
        {

            log INFO "  sendMsg to $($subscriber): STARTED"

            $message = New-Object System.Net.Mail.MailMessage
            $message.subject = "lunchBot $subDate"
            $message.body = $msg
            $message.to.add($subscriber)
            $message.from = $Username

            $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
            $smtp.EnableSSL = $true
            $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
            $smtp.send($message)

        }
        Catch [Exception]
        {
            $exMsg = $_.Exception.Message
            HandleException "sendMsg" $exMsg
        }

        log INFO "  sendMsg to $($subscriber): COMPLETED"
    }
}

############################## script entry point
log INFO "script: STARTED"
clear

$msg = getMsg $locations

if ($msg)
{
    sendMsg $subscribers $msg
}
else
{
    # this will happen on weekends, holidays, or any other edge case where there are no restaurants found on a menu page with today's date
    log INFO 'getMsg returned $null, no messages will be sent.'
}
    
log INFO "script: COMPLETED, exiting"
