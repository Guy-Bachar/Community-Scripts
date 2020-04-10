<#  
.SYNOPSIS  
	Shows a user contact list

.NOTES  
  Version      				: 0.2
  Rights Required			: Local admin on server and SQL (Run PowerShell as Administrator)
  Lync Version				: 2013 (tested on March 2014 Updated)
  Author       				: Guy Bachar
  Twitter/Blog	            : @GuyBachar, http://guybachar.us

.SOURCES
    SQL Query: http://social.technet.microsoft.com/Forums/en-US/2ba502b7-4aef-43c2-adc1-ad601597d23f/contacts-delete-lync-contacts-for-a-user-on-server-side?forum=ocsmanagement
    
.EXAMPLE
	 .\Get-CsUserContactList.ps1

.VERSIONS
    0.1 - Initial Configuration
    0.2 - Fixed SQL Server Connections details
	
#>


Clear-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Get User Cotact List" -ForegroundColor Green
Write-Host "Version: 0.2" -ForegroundColor Green
Write-Host "Author: Guy Bachar | @GuyBachar | http://guybachar.us" -ForegroundColor Green
Write-host
$Date = Get-Date -DisplayHint DateTime
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Data collected:" , $Date -ForegroundColor Yellow
Write-host

#Import Lync Module
Import-Module Lync

#Verify if the Script is running under Admin privliges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Write-Host 
    Break
}

#Create empty variable that will contain the user contacts list
$overallrecords = $null

#Input for User name
Do
{
    $Error.Clear()
    $UserName = Read-Host "Please enter user SIP address or username"
	try 
	{
	$LyncUser = Get-CsUser -Identity $UserName -ErrorAction Stop
	}
	Catch [Exception]
	{
	Write-Host "User " -nonewline; Write-Host $UserName -foregroundcolor red -nonewline; " does not exist or cannot be contacted. Please try again."
Write-Host
	}
}
until (($LyncUser -ne $Null) -and (!$Error))

$ServerName = $LyncUser.RegistrarPool.FriendlyName

#Defined Connection String
$connstring = "server=$ServerName\rtclocal;database=rtc;`
    trusted_connection=true;"

#Removeing SIP: from the User SIP Address
$SipAddress = $LyncUser.SipAddress -replace "sip:", ""

#Define SQL Command     
$command = New-Object System.Data.SqlClient.SqlCommand
$command.CommandText = "SELECT [rtc].[dbo].[Contact].[OwnerId], [rtc].[dbo].[Contact].[BuddyId],[rtc].[dbo].[ContactGroupAssoc].[GroupNumber], [rtc].[dbo].[Resource].[UserAtHost] `
                        FROM [rtc].[dbo].[Contact] `
                        INNER JOIN [rtc].[dbo].[Resource] ON [rtc].[dbo].[Contact].[BuddyId] = [rtc].[dbo].[Resource].[ResourceId] `
                        LEFT OUTER JOIN [rtc].[dbo].[ContactGroupAssoc] ON ([rtc].[dbo].[Contact].[BuddyId] = [rtc].[dbo].[ContactGroupAssoc].[BuddyId] AND [rtc].[dbo].[Contact].[OwnerId] = [rtc].[dbo].[ContactGroupAssoc].[OwnerId]) `
                        WHERE [rtc].[dbo].[Contact].[OwnerId] =(SELECT ResourceId from [rtc].[dbo].[Resource] WHERE [rtc].[dbo].[Resource].[UserAtHost] = '$SipAddress') `
                        ORDER BY [rtc].[dbo].[ContactGroupAssoc].[GroupNumber] DESC "
		
#Make the connection to Server     
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connstring
$connection.Open()
$command.Connection = $connection
     
#Get the results
$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
$sqladapter.SelectCommand = $command
$results = New-Object System.Data.Dataset     
$recordcount=$sqladapter.Fill($results) 
     
#Close the connection
$connection.Close()
     
#Append the results to the reuslts from the previous servers
$overallrecords = $overallrecords + $Results.Tables[0] 

$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\$SIP-"+$FileDate+".csv"

#Getting Output Options
Write-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "How would you like to see the results?" -ForegroundColor Yellow
Write-Host "1) Grid View"
Write-Host "2) CSV File"
Write-Host "3) Screen"
Write-Host
$ExportOptions = Read-Host "Please Enter your choice"
switch ($ExportOptions) 
    { 
        1 {"You chose GridView"} 
        2 {"You chose CSV"}
        3 {"You chose Screen"}
	   default {"A Non valid selection was chosen, results will output to screen."}
    }

if ($ExportOptions -eq 1)
{ $overallrecords | Select * | Out-GridView}
elseif ($ExportOptions -eq 2)
{
    $overallrecords | Export-Csv -Path $ServicesFileName -NoTypeInformation
    Invoke-Expression "explorer '/select,$ServicesFileName'"
}
else
{ $overallrecords }
