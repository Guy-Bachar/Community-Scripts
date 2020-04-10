<#
.SYNOPSIS
    The script output will will be the amount of users assigned to every policy.

.DESCRIPTION
    	
	This script will count the amount of users assigned to Lync Policies (Client, Voice, Conferencing, Hosted Voice Mail and Mobility).
    Using this script you can find unused policies or misconfigured ones.
	
.NOTES
    File Name: CountUsersPerPolicy.ps1
	Version: 0.1
	Last Update: 18-May-2014
    Author: Guy Bachar, @GuyBachar, http://guybachar.us"
    The script are provided “AS IS” with no guarantees, no warranties, USE ON YOUR OWN RISK.    

.SOURCE
    Concept taken from - http://blogs.technet.com/b/csps/archive/2010/06/06/scriptnumberofusersassignedtopolicies.aspx
    Write-Host Alignment - Taken from Pat Richard Script, Get-CsConnections.ps1

.WHATSNEW
    0.1 - Added 5 initial policies for counter
#> 

Clear-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Count Users Per Policy" -ForegroundColor Green
Write-Host "Version: 0.1" -ForegroundColor Green
Write-Host 
Write-Host "Authors:" -ForegroundColor Green
Write-Host
Write-Host "Guy Bachar        @GuyBachar     http://guybachar.us" -ForegroundColor Green
Write-host
$Date = Get-Date -DisplayHint DateTime
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Data collected:" , $Date -ForegroundColor Yellow
Write-host


# Declare Varibles
$VoicePolicies = @()
$ConfPolicies = @()
$ClientPolicies = @()
$HostedVoicemailPolicies = @()
$MobilityPolicies = @()
$DialPlans = @()
$ExternalAccessPolicies = @()

# Function to remove TAG: from policies name
function RemoveTag
    {
    param($PolicyToRemoveTagFrom)
    $Identities = @()
    foreach ($i in $PolicyToRemoveTagFrom)
        {
            $x = $i.Identity
            $x = $x -replace "Tag:",""
            #$x = $x -replace "Site:",""
            $identities += $x
        }
        #Write-Host $Identities
        return $Identities
    }


$VoicePolicies = RemoveTag(Get-CsVoicePolicy | Select-Object Identity)
$ConfPolicies = RemoveTag(Get-CsConferencingPolicy | Select-Object Identity)
$ClientPolicies = RemoveTag(Get-CsClientPolicy | Select-Object Identity)
$HostedVoicemailPolicies = RemoveTag(Get-CsHostedVoicemailPolicy | Select-Object Identity)
$MobilityPolicies = RemoveTag(Get-CsMobilityPolicy | Select-Object Identity)
$DialPlans = RemoveTag(Get-CsDialPlan | Select-Object Identity)
$ExternalAccessPolicies = RemoveTag(Get-CsExternalAccessPolicy | Select-Object Identity)


Write-Host ("{0,-41}{1,15}" -f "Voice Policy", "Total") -ForegroundColor Cyan 
foreach ($policy in $VoicePolicies)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.VoicePolicy -eq $null}).Count)            
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.VoicePolicy -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {VoicePolicy -eq $policy}).Count)            
        }
    }


Write-Host ("{0,-41}{1,15}" -f "`nConferencing Policy", "Total") -ForegroundColor Cyan 
foreach ($policy in $Confpolicies)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.ConferencingPolicy -eq $null}).Count)
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.ConferencingPolicy -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {ConferencingPolicy -eq $policy}).Count)            
        }
    }


Write-Host ("{0,-41}{1,15}" -f "`nClient Policy", "Total") -ForegroundColor Cyan 
foreach ($policy in $ClientPolicies)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.ClientPolicy -eq $null}).Count)
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.ClientPolicy -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {ClientPolicy -eq $policy}).Count)            
        }
    }


Write-Host ("{0,-41}{1,15}" -f "`nHosted Voice Mail Policy", "Total") -ForegroundColor Cyan 
foreach ($policy in $HostedVoicemailPolicies)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.HostedVoicemailPolicy -eq $null}).Count)
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.HostedVoicemailPolicy -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {HostedVoicemailPolicy -eq $policy}).Count)            
        }
    }


Write-Host ("{0,-41}{1,15}" -f "`nMobility Policy", "Total") -ForegroundColor Cyan 
foreach ($policy in $MobilityPolicies)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.MobilityPolicy -eq $null}).Count)
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.MobilityPolicy -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {MobilityPolicy -eq $policy}).Count)            
        }
    }


Write-Host ("{0,-41}{1,15}" -f "`nExternal Access Policies", "Total") -ForegroundColor Cyan 
foreach ($policy in $ExternalAccessPolicies)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.ExternalAccessPolicy -eq $null}).Count)            
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.ExternalAccessPolicy -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {ExternalAccessPolicy -eq $policy}).Count)            
        }
    }

Write-Host ("{0,-41}{1,15}" -f "`nDial Plans", "Total") -ForegroundColor Cyan 
foreach ($policy in $DialPlans)
    {      
        if ($policy -match "Global")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.DialPlan -eq $null}).Count)            
        }
        elseif ($policy -like "Site:*")
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser | Where-Object {$_.DialPlan -like $policy}).Count)            
        }
        else
        {
            Write-Host ("{0,-41}{1,15}" -f $policy, (Get-CsUser -Filter {DialPlan -eq $policy}).Count)            
        }
    }