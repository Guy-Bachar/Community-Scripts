<#
.SYNOPSIS
    The script output will be the different settings two Lync users have based on their assigned policies.

.DESCRIPTION
    	
	This script will compare between two Lync users to find the differences in setting between the users based on their assigned policies.
    In some cases we are required to verify why user A is capable of certain functionality user B is not. In most of the time we are verifying the policies manually by using the Lync Control Panel.
    This script allow you to get as input two users and it will output all the differences between them based on their policies.
	
.NOTES
    File Name: CompareLyncUsersPolicies.ps1
	Version: 0.5
	Last Update: 18-May-2014
    Author: Guy Bachar, @GuyBachar, http://guybachar.us"
    Author: Yoav Barzilay, @y0av, http://y0av.me/"
    The script are provided “AS IS” with no guarantees, no warranties, USE ON YOUR OWN RISK.    

.WHATSNEW
    0.1 - Added Main policies for comparison: Conference and Client Policy
    0.2 - Added More policies Comparison: Dial Plan, Voice, HostedVoiceMail, Mobility
    0.3 - Fixed HostedVoiceMail Policy error, added colorized warnings
    0.4 - Added HTML Export Report
    0.5 - Fixed Display Names for Users, Added Admin Privileges verification
#> 

Clear-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Compare Lync Policies" -ForegroundColor Green
Write-Host "Version: 0.5" -ForegroundColor Green
Write-Host 
Write-Host "Authors:" -ForegroundColor Green
Write-Host
Write-Host " Guy Bachar        @GuyBachar     http://guybachar.us" -ForegroundColor Green
Write-Host " Yoav Barzilay     @y0avb         http://y0av.me" -ForegroundColor Green
Write-host
$Date = Get-Date -DisplayHint DateTime
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host
Write-Host "Data collected:" , $Date -ForegroundColor Yellow
Write-Host
Write-host

#Variables
$LyncUser1 = $null
$LyncUser2 = $null
$ExportToHtml = $null
$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\CompareLyncUsersPoliciesReport-"+$FileDate+".htm"

Import-Module Lync

#Verify if the Script is running under Admin privliges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Write-Host 
    Break
}

# Input for User 1
Do
{
    $Error.Clear()
    $InputUser1 = Read-Host "First user's SIP address or username"
	try 
	{
	$LyncUser1 = Get-CsUser -Identity $InputUser1 -ErrorAction Stop
	}
	Catch [Exception]
	{
	Write-Host "User " -nonewline; Write-Host $InputUser1 -foregroundcolor red -nonewline; " does not exist or cannot be contacted. Please try again."
Write-Host
	}
}
until (($LyncUser1 -ne $Null) -and (!$Error))

# Input for User 2
Do
{
    $Error.Clear()	
    $InputUser2 = Read-Host "Second user's SIP address or username"
	try 
	{
	$LyncUser2 = Get-CsUser -Identity $InputUser2 -ErrorAction Stop 
	}
	Catch [Exception]
	{
	Write-Host "User " -nonewline; Write-Host $InputUser2 -foregroundcolor red -nonewline; " does not exist or cannot be contacted. Please try again."
Write-Host
	}
}
until (($LyncUser1 -ne $Null) -and (!$Error))


####  Getting User Input  ####
Write-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Would you like to export the report to HTML?" -ForegroundColor Yellow
Write-Host "1) Yes"
Write-Host "2) No"
Write-Host
$ExportToHtml = Read-Host "Please Enter your choice"
switch ($ExportToHtml) 
    { 
        1 {"You chose to Export to HTML"} 
        2 {"You chose not to Export to HTML"}
	   default {"A Non valid selection was chosen, results will output to screen."}
    }

####  Building HTML File  ####
Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Users Policies Comparison Report</title>'
add-content $fileName '<STYLE TYPE="text/css">'
add-content $fileName  "<!--"
add-content $fileName  "td {"
add-content $fileName  "font-family: Tahoma;"
add-content $fileName  "font-size: 11px;"
add-content $fileName  "border-top: 1px solid #999999;"
add-content $fileName  "border-right: 1px solid #999999;"
add-content $fileName  "border-bottom: 1px solid #999999;"
add-content $fileName  "border-left: 1px solid #999999;"
add-content $fileName  "padding-top: 0px;"
add-content $fileName  "padding-right: 0px;"
add-content $fileName  "padding-bottom: 0px;"
add-content $fileName  "padding-left: 0px;"
add-content $fileName  "}"
add-content $fileName  "body {"
add-content $fileName  "margin-left: 5px;"
add-content $fileName  "margin-top: 5px;"
add-content $fileName  "margin-right: 0px;"
add-content $fileName  "margin-bottom: 10px;"
add-content $fileName  ""
add-content $fileName  "table {"
add-content $fileName  "border: thin solid #000000;"
add-content $fileName  "}"
add-content $fileName  "-->"
add-content $fileName  "</style>"
add-content $fileName  "</head>"
add-content $fileName  "<body>"
add-content $fileName  "<table width='100%'>"
add-content $fileName  "<tr bgcolor='#CCCCCC'>"
add-content $fileName  "<td colspan='7' height='25' align='center'>"
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Lync Policy Comparison Report - $date</strong></font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName,$User1Name1,$User2Name2)
Add-Content $fileName "<tr bgcolor=#CCCCCC>"
Add-Content $fileName "<td width='20%' align='center'>Attribute</td>"
Add-Content $fileName "<td width='40%' align='center'>$User1Name1</td>"
Add-Content $fileName "<td width='40%' align='center'>$User2Name2</td>"
Add-Content $fileName "</tr>"
}

Function writeHtmlFooter
{
param($fileName)
Add-Content $fileName "</body>"
Add-Content $fileName "</html>"
}

Function writeServiceInfo
{
param($fileName,$PolicyAtt,$User1Att,$User2Att)
 if ($User1Att -eq $User2Att)
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$PolicyAtt</td>"
 Add-Content $fileName "<td align=center>$User1Att</td>"
 Add-Content $fileName "<td align=center>$User2Att</td>"
 Add-Content $fileName "</tr>"
 }
 else
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$PolicyAtt</td>"
 Add-Content $fileName "<td bgcolor='#FBB917' align=center>$User1Att</td>"
 Add-Content $fileName "<td bgcolor='#FBB917' align=center>$User2Att</td>"
 Add-Content $fileName "</tr>"
 }
}
####  Closing HTML File  ####

Function GetUserConferencePolicy
{ param($UserName)
if ($UserName.ConferencingPolicy.FriendlyName -eq $null)
    { $LyncUserConfernecePolicyProperties = @((Get-CsConferencingPolicy -Identity Global).psobject.properties) }
else{ $LyncUserConfernecePolicyProperties = @((Get-CsConferencingPolicy -Identity $UserName.ConferencingPolicy.FriendlyName).psobject.properties) }
return $LyncUserConfernecePolicyProperties }

Function GetUserClientPolicy 
{ param($UserName)
if ($UserName.ClientPolicy.FriendlyName -eq $null)
    { $LyncUserClientPolicyProperties = @((Get-CsClientPolicy -Identity Global).psobject.properties) }
else{ $LyncUserClientPolicyProperties = @((Get-CsClientPolicy -Identity $UserName.ClientPolicy.FriendlyName).psobject.properties) }
return $LyncUserClientPolicyProperties }

Function GetUserVoicePolicy 
{ param($UserName)
if ($UserName.VoicePolicy.FriendlyName -eq $null)
    { $LyncUserVoicePolicyProperties = @((Get-CsVoicePolicy -Identity Global).psobject.properties) }
else{ $LyncUserVoicePolicyProperties = @((Get-CsVoicePolicy -Identity $UserName.VoicePolicy.FriendlyName).psobject.properties) }
return $LyncUserVoicePolicyProperties }

Function GetUserDialPlan 
{ param($UserName)
if ($UserName.DialPlan.FriendlyName -eq $null)
    { $LyncUserDialPlanProperties = @((Get-CsDialPlan -Identity Global).psobject.properties) }
else{ $LyncUserDialPlanProperties = @((Get-CsDialPlan -Identity $UserName.DialPlan.FriendlyName).psobject.properties) }
return $LyncUserDialPlanProperties }

Function GetUserHostedVoicemailPolicy 
{ param($UserName)
if ($UserName.HostedVoicemailPolicy.FriendlyName -eq $null)
    { $LyncUserHostedVoicemailPolicyProperties = @((Get-CsHostedVoicemailPolicy -Identity Global).psobject.properties) }
else{ $LyncUserHostedVoicemailPolicyProperties = @((Get-CsHostedVoicemailPolicy -Identity $UserName.HostedVoicemailPolicy.FriendlyName).psobject.properties) }
return $LyncUserHostedVoicemailPolicyProperties }


$LyncUser1ConfernecePolicyProperties = GetUserConferencePolicy($LyncUser1)
$LyncUser2ConfernecePolicyProperties = GetUserConferencePolicy($LyncUser2)
$LyncUser1ClientPolicyProperties = GetUserClientPolicy($LyncUser1)
$LyncUser2ClientPolicyProperties = GetUserClientPolicy($LyncUser2)
$LyncUser1VoicePolicyProperties = GetUserVoicePolicy($LyncUser1)
$LyncUser2VoicePolicyProperties = GetUserVoicePolicy($LyncUser2)
$LyncUser1DialPlanPolicyProperties = GetUserDialPlan($LyncUser1)
$LyncUser2DialPlanPolicyProperties = GetUserDialPlan($LyncUser2)
$LyncUser1HostedVoicemailPolicyProperties = GetUserHostedVoicemailPolicy($LyncUser1)
$LyncUser2HostedVoicemailPolicyProperties = GetUserHostedVoicemailPolicy($LyncUser2)

If ($ExportToHtml -eq 1) {
	#### Adding Content to HTML ####
    $User1Name = $LyncUser1.DisplayName.ToString()
    $User2Name = $LyncUser2.DisplayName.ToString()
	writeHtmlHeader $ServicesFileName

	    Add-Content $ServicesFileName "<table width='100%'><tbody>"
	    Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
	    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> Conference Policy </strong></font></td>"
	    Add-Content $ServicesFileName "</tr>"

	    writeTableHeader $ServicesFileName $User1Name $User2Name
	    for ($i=0; $i -lt ($LyncUser1ConfernecePolicyProperties.count-1); $i++)
	    {
	    writeServiceInfo $ServicesFileName $LyncUser1ConfernecePolicyProperties[$i].Name $LyncUser1ConfernecePolicyProperties[$i].Value $LyncUser2ConfernecePolicyProperties[$i].Value
	    }

	    Add-Content $ServicesFileName "<table width='100%'><tbody>"
	    Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
	    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> Client Policy </strong></font></td>"
	    Add-Content $ServicesFileName "</tr>"
	    writeTableHeader $ServicesFileName $User1Name $User2Name
	    for ($i=0; $i -lt ($LyncUser1ClientPolicyProperties.count-1); $i++)
	    {
	    writeServiceInfo $ServicesFileName $LyncUser1ClientPolicyProperties[$i].Name $LyncUser1ClientPolicyProperties[$i].Value $LyncUser2ClientPolicyProperties[$i].Value
	    }
	    
	    Add-Content $ServicesFileName "<table width='100%'><tbody>"
	    Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
	    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> Voice Policy </strong></font></td>"
	    Add-Content $ServicesFileName "</tr>"
	    writeTableHeader $ServicesFileName $User1Name $User2Name           
	    for ($i=0; $i -lt ($LyncUser1VoicePolicyProperties.count-1); $i++)
	    {
	    writeServiceInfo $ServicesFileName $LyncUser1VoicePolicyProperties[$i].Name $LyncUser1VoicePolicyProperties[$i].Value $LyncUser2VoicePolicyProperties[$i].Value
	    }
	    
	    Add-Content $ServicesFileName "<table width='100%'><tbody>"
	    Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
	    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> Dial Plan </strong></font></td>"
	    Add-Content $ServicesFileName "</tr>"
	    writeTableHeader $ServicesFileName $User1Name $User2Name         
	    for ($i=0; $i -lt ($LyncUser1DialPlanPolicyProperties.count-1); $i++)
	    {
	    writeServiceInfo $ServicesFileName $LyncUser1DialPlanPolicyProperties[$i].Name $LyncUser1DialPlanPolicyProperties[$i].Value $LyncUser2DialPlanPolicyProperties[$i].Value
	    }

	    Add-Content $ServicesFileName "<table width='100%'><tbody>"
	    Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
	    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> Hosted Voicemail Policy </strong></font></td>"
	    Add-Content $ServicesFileName "</tr>"
	    writeTableHeader $ServicesFileName $User1Name $User2Name
	    for ($i=0; $i -lt ($LyncUser1HostedVoicemailPolicyProperties.count-1); $i++)
	    {
	    writeServiceInfo $ServicesFileName $LyncUser1HostedVoicemailPolicyProperties[$i].Name $LyncUser1HostedVoicemailPolicyProperties[$i].Value $LyncUser2HostedVoicemailPolicyProperties[$i].Value
	    }
        Add-Content $ServicesFileName "</table>" 
        writeHtmlFooter $ServicesFileName
        Invoke-Item $ServicesFileName
    }
#### Closing HTML File ####
else 
{
    #### Writing Output to Screen ####
Write-Host
Write-Host "Conferencing Policy Comparison:" -ForegroundColor Green
if(($LyncUser1.ConferencingPolicy) -eq ($LyncUser2.ConferencingPolicy))
     { Write-Host "Both Conference Policies match" -ForegroundColor Yellow }
else {
Write-Host "Conference Policies don't match" -ForegroundColor Red
        for($i=0; $i -lt ($LyncUser1ConfernecePolicyProperties.count-1); $i++)
         {
            # create custom object if a property value is different
            if("$($LyncUser1ConfernecePolicyProperties[$i].value)" -ne "$($LyncUser2ConfernecePolicyProperties[$i].value)")
            {
                New-Object PSObject | 
                Add-Member NoteProperty -Name Property -Value $LyncUser1ConfernecePolicyProperties[$i].Name -PassThru | 
                Add-Member NoteProperty -Name $LyncUser1.DisplayName -Value $LyncUser1ConfernecePolicyProperties[$i].value -PassThru | 
                Add-Member NoteProperty -Name $LyncUser2.DisplayName -Value $LyncUser2ConfernecePolicyProperties[$i].value -PassThru
             }
         }
       }


Write-Host
Write-Host "Client Policy Comparison:" -ForegroundColor Green
if(($LyncUser1.ClientPolicy) -eq ($LyncUser2.ClientPolicy))
     { Write-Host "Both Client Policies match" -ForegroundColor Yellow }
else {
Write-Host "Client Policies don't match" -ForegroundColor Red
        for($i=0; $i -lt ($LyncUser1ClientPolicyProperties.count-1); $i++)
         {
            # create custom object if a property value is different
            if("$($LyncUser1ClientPolicyProperties[$i].value)" -ne "$($LyncUser2ClientPolicyProperties[$i].value)")
             {
                New-Object PSObject | 
                Add-Member NoteProperty -Name Property -Value $LyncUser1ClientPolicyProperties[$i].Name -PassThru | 
                Add-Member NoteProperty -Name $LyncUser1.DisplayName -Value $LyncUser1ClientPolicyProperties[$i].value -PassThru | 
                Add-Member NoteProperty -Name $LyncUser2.DisplayName -Value $LyncUser2ClientPolicyProperties[$i].value -PassThru
             }
         }
      }


Write-Host
Write-Host "Voice Policy Comparison:" -ForegroundColor Green
if(($LyncUser1.VoicePolicy) -eq ($LyncUser2.VoicePolicy))
     { Write-Host "Both Voice Policies match" -ForegroundColor Yellow }
else {
Write-Host "Voice Policies don't match" -ForegroundColor Red
        for($i=0; $i -lt ($LyncUser1VoicePolicyProperties.count-1); $i++)
         {
            # create custom object if a property value is different
            if("$($LyncUser1VoicePolicyProperties[$i].value)" -ne "$($LyncUser2VoicePolicyProperties[$i].value)")
             {
                New-Object PSObject | 
                Add-Member NoteProperty -Name Property -Value $LyncUser1VoicePolicyProperties[$i].Name -PassThru | 
                Add-Member NoteProperty -Name $LyncUser1.DisplayName -Value $LyncUser1VoicePolicyProperties[$i].value -PassThru | 
                Add-Member NoteProperty -Name $LyncUser2.DisplayName -Value $LyncUser2VoicePolicyProperties[$i].value -PassThru
             }
         }
      }


Write-Host
Write-Host "Dial Plan Comparison:" -ForegroundColor Green
if(($LyncUser1.DialPlan) -eq ($LyncUser2.DialPlan))
     { Write-Host "Both Dial Plans match" -ForegroundColor Yellow }
else {
Write-Host "Dial Plans don't match" -ForegroundColor Red
             for($i=0; $i -lt ($LyncUser1DialPlanPolicyProperties.count); $i++)
         {
            # create custom object if a property value is different
            if("$($LyncUser1DialPlanPolicyProperties[$i].value)" -ne "$($LyncUser2DialPlanPolicyProperties[$i].value)")
             {
                New-Object PSObject | 
                Add-Member NoteProperty -Name Property -Value $LyncUser1DialPlanPolicyProperties[$i].Name -PassThru | 
                Add-Member NoteProperty -Name $LyncUser1.DisplayName -Value $LyncUser1DialPlanPolicyProperties[$i].value -PassThru | 
                Add-Member NoteProperty -Name $LyncUser2.DisplayName -Value $LyncUser2DialPlanPolicyProperties[$i].value -PassThru
             }
         }
       }


Write-Host
Write-Host "Hosted Voice Mail Policy Comparison:" -ForegroundColor Green
if(($LyncUser1.HostedVoicemailPolicy) -eq ($LyncUser2.HostedVoicemailPolicy))
     { Write-Host "Both Hosted Voicemail Policies match" -ForegroundColor Yellow }
else {
Write-Host "Hosted Voicemail Policies don't match" -ForegroundColor Red
        for($i=0; $i -lt ($LyncUser1HostedVoicemailPolicyProperties.count-1); $i++)
         {
            # create custom object if a property value is different
            if("$($LyncUser1HostedVoicemailPolicyProperties[$i].value)" -ne "$($LyncUser2HostedVoicemailPolicyProperties[$i].value)")
             {
                New-Object PSObject | 
                Add-Member NoteProperty -Name Property -Value $LyncUser1HostedVoicemailPolicyProperties[$i].Name -PassThru | 
                Add-Member NoteProperty -Name $LyncUser1.DisplayName -Value $LyncUser1HostedVoicemailPolicyProperties[$i].value -PassThru | 
                Add-Member NoteProperty -Name $LyncUser2.DisplayName -Value $LyncUser2HostedVoicemailPolicyProperties[$i].value -PassThru
             }
         }
       }

}