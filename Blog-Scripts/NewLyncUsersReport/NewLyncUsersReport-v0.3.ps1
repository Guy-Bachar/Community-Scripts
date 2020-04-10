<#  
.SYNOPSIS  
    This script will generate an HTML report which include all the users that were enabled for Lync within X amount of days the user provide as an input

.NOTES  
    Version                   : 0.3
    Rights Required           : Local admin
    Lync Version              : 2013 (tested on August 2014 CU5 Update)
    Authors                   : Guy Bachar
    Last Update               : 30-November-2014
    Twitter/Blog              : @GuyBachar, http://guybachar.us


.VERSION
    0.1 - Initial Version for reporting New Lync Users
    0.2 - Added SQL integration to get Lync Creation date
    0.3 - Added Support for parameters and scheduled tasks and script runtime optimizations
       
#>

Param (
	[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory=$true, Position=0)]
	[string] $PoolFQDN,
    
    [Parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory=$False, Position=1)]  
    #If not inserted, the default is one week back
    [String]$HowManyDaysBack = 7
)


#region Script Information
Clear-Host
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "New Lync Users Report" -ForegroundColor Green
Write-Host "Version: 0.3" -ForegroundColor Green
Write-Host 
Write-Host "Authors:" -ForegroundColor Green
Write-Host " Guy Bachar       | @GuyBachar        | http://guybachar.us" -ForegroundColor Green
Write-host
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
#endregion

#region Verifying Administrator Elevation
Write-Host Verifying User permissions... -ForegroundColor Yellow
#Verify if the Script is running under Admin privileges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
  [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
  Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!"
  Write-Host 
  Break
}
#endregion

#region Import Lync Module
Write-Host
Write-Host "Please wait while the Lync PowerShell Module is loading..." -ForegroundColor Yellow
  if(-not (Get-Module -Name "Lync")){
    if(Get-Module -Name "Lync" -ListAvailable){
      Import-Module -Name "Lync";
      Write-Host "Loading Lync Module" -ForegroundColor Yellow
    }
    else{
      Write-Host Lync Module does not exist on this computer, please verify the Lync Admin tools installed -ForegroundColor Red
      exit;   
    }    
  }
Write-Host    
Write-Host "Done!" -ForegroundColor Green
#endregion

#region Setting Variables
$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LyncNewUsersReport-"+$FileDate+".htm"
New-Item -ItemType file $ServicesFileName -Force
$DaysBack = (Get-Date).AddDays(-$HowManyDaysBack)
$SQLDaysBack = "{0:yyyy-MM-dd}" -f (Get-Date).AddDays(-$HowManyDaysBack)
$StopWatch = New-Object System.Diagnostics.Stopwatch 
$StopWatch.Start() 
#endregion

function ConvertTo-SID
{
    param([byte[]]$ObjectSid)

    $sid = New-Object System.Security.Principal.SecurityIdentifier $ObjectSid,0 
    $sid.Translate([System.Security.Principal.NTAccount]).Value
}

#region Loop Through Front-End Pool
$strResults = $null
$CSPool = Get-CSPool $PoolFQDN

Foreach ($Computer in $CSPool.Computers){

    #Get Computer Name
    $ComputerName = $Computer.Split(".")[0]

    #Connection String
    $sqlConnString = "server=$ComputerName\rtclocal;database=rtc;trusted_connection=true;"

    #SQL Command     
    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand

    $sqlCommand.CommandText = "SELECT [ResourceId],[UserPrincipalName],[AdUserSid],[InsertTime],[UpdateTime],[Enabled],[SmtpUserAtHost],[AdDisplayName] `
    FROM [rtc].[dbo].[ResourceDirectory] `
      WHERE AdUserSid IS NOT NULL AND InsertTime > Convert(datetime, '$SQLDaysBack')"

    #Connect to SQL    
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $sqlConnection.ConnectionString = $sqlConnString
    $sqlConnection.Open()
    $sqlCommand.Connection = $sqlConnection
     
    #Query Server
    $sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $sqlAdapter.SelectCommand = $sqlCommand
    $results = New-Object System.Data.Dataset     
    $recordcount=$sqladapter.Fill($results) 
     
    #Close DB Connection
    $sqlConnection.Close()
     
    #Append All strResults
    $strResults = $strResults + $Results.Tables[0] 
}

$strResultsFixed = $strResults | Select-Object *
$strResultsFixed | ForEach-Object {
   $_.AdUserSid = ConvertTo-SID -ObjectSid $_.AdUserSid
}
$strResultsFixed = $strResultsFixed | Sort-Object -Property AdUserSid -Unique
#endregion

#### Building HTML File ####
Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Users Report</title>'
add-content $fileName '<STYLE TYPE="text/css">'
add-content $fileName  "<!--"
add-content $fileName  "td {"
add-content $fileName  "font-family: Segoe UI;"
add-content $fileName  "font-size: 11px;"
add-content $fileName  "border-top: 1px solid #1E90FF;"
add-content $fileName  "border-right: 1px solid #1E90FF;"
add-content $fileName  "border-bottom: 1px solid #1E90FF;"
add-content $fileName  "border-left: 1px solid #1E90FF;"
add-content $fileName  "padding-top: 0px;"
add-content $fileName  "padding-right: 0px;"
add-content $fileName  "padding-bottom: 0px;"
add-content $fileName  "padding-left: 0px;"
add-content $fileName  "}"
add-content $fileName  "body {"
add-content $fileName  "margin-left: 5px;"
add-content $fileName  "margin-top: 5px;"
add-content $fileName  "margin-right: 0px;"
add-content $fileName  "margin-bottom: 10px;"
add-content $fileName  ""
add-content $fileName  "table {"
add-content $fileName  "border: thin solid #000000;"
add-content $fileName  "}"
add-content $fileName  "-->"
add-content $fileName  "</style>"
add-content $fileName  "</head>"
add-content $fileName  "<body>"
add-content $fileName  "<table width='100%'>"
add-content $fileName  "<tr bgcolor='#336699 '>"
add-content $fileName  "<td colspan='10' height='25' align='center'>"
add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'>$date - New Lync Users Report (The following users were enabled for Lync within the last $HowManyDaysBack days)</font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr bgcolor=#0099CC>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Display Name</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>SamAccountName</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>SIP Address</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Line URI</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Registrar</font></td>"
Add-Content $fileName "<td width='8%' align='center'><font color=#FFFFFF>EV Enabled</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Conference Policy</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>External Policy</font></td>"
Add-Content $fileName "<td width='11%' align='center'><font color=#FFFFFF>Creation Date</font></td>"
Add-Content $fileName "<td width='11%' align='center'><font color=#FFFFFF>Update Date</font></td>"
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
param($fileName,$DisplayName,$SamACcountName,$SIP,$Registrar,$EVE,$LineURI,$WC,$Conf,$External,$WU)
$DiffDay = ((Get-Date)-$WC).Days
$DiffDay2 = ((Get-Date)-$WU).Days
Add-Content $fileName "<tr>"
Add-Content $fileName "<td width='10%' align='Center'>$DisplayName</td>"
Add-Content $fileName "<td width='10%' align='Center'>$SamACcountName</td>"
Add-Content $fileName "<td width='10%' align='Center'>$SIP</td>"
Add-Content $fileName "<td width='10%' align='Center'>$LineURI</td>"
Add-Content $fileName "<td width='10%' align='Center'>$Registrar</td>"
Add-Content $fileName "<td width='8%' align='Center'>$EVE</td>"
Add-Content $fileName "<td width='10%' align='Center'>$Conf</td>"
Add-Content $fileName "<td width='10%' align='Center'>$External</td>"
Add-Content $fileName "<td width='11%' align='Center'>$WC - ($diffday)</td>"
Add-Content $fileName "<td width='11%' align='Center'>$WU - ($diffday2)</td>"
Add-Content $fileName "</tr>"
}

Function sendEmail
{ param($from,$to,$subject,$smtphost,$htmlFileName)
$body = Get-Content $htmlFileName
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
$msg.isBodyhtml = $true
$smtp.send($msg)
}

# Main Script
$UserLoopCount = 0
$LyncUsersList = @()
:LyncUserProcessin foreach ($rowobject in $strResultsFixed)
{
    $PercentComplete = [Math]::Round(($UserLoopCount++ / $strResultsFixed.Count * 100),1)
    $CurrentUser = $rowobject.AdUserSid
    Write-Progress -Activity ("User data gathering in progress on Lync Pool: $PoolFQDN") `
    -PercentComplete $PercentComplete -Status "$PercentComplete% Complete" -CurrentOperation "Current Lync User: $CurrentUser" 
    $tempusr = Get-CsUser -Identity $rowobject.AdUserSid -ErrorAction SilentlyContinue
    $LyncUsersList += $tempusr | Select-Object DisplayName,SamAccountName,SipAddress,RegistrarPool,EnterpriseVoiceEnabled,LineURI,ConferencingPolicy,ExternalAccessPolicy, `
    @{Name="WhenCreated";Expression={$rowobject.InsertTime}},@{Name="WhenChanged";Expression={$rowobject.UpdateTime}}
}
Write-Progress -Activity "User data gathering in progress on Lync Server:" -Completed -Status "Completed" 

writeHtmlHeader $ServicesFileName
Add-Content $ServicesFileName "<table width='100%'><tbody>"
WriteTableHeader $ServicesFileName

$LyncUsersList = $LyncUsersList | Sort-Object -Property WhenCreated -Descending

foreach ($User in $LyncUsersList)
{       
        foreach ($item in $User)
        {
            writeServiceInfo $ServicesFileName $item.DisplayName $item.SamAccountName $item.SipAddress $item.RegistrarPool $item.EnterpriseVoiceEnabled `
            $item.LineURI $item.WhenCreated $item.ConferencingPolicy $item.ExternalAccessPolicy $item.WhenChanged
        }
}
Add-Content $ServicesFileName "</table>"
writeHtmlFooter $ServicesFileName

### Configuring Email Parameters
#sendEmail from@domain.com to@domain.com "Lync Users Report - $Date" SMTP_SERVER $ServicesFileName

#Closing HTML
writeHtmlFooter $ServicesFileName
$StopWatch.Stop() 
$ElapsedTime = $StopWatch.Elapsed 
Write-Host "`nThe script ran for " $ElapsedTime.Hours "hours," $ElapsedTime.Minutes "minutes, and" $ElapsedTime.Seconds "seconds" -ForegroundColor Yellow
Write-Host "`nThe File was generated at the following location: $ServicesFileName `n`nOpenning file..." -ForegroundColor Yellow
Invoke-Item $ServicesFileName
