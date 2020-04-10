﻿<#  
.SYNOPSIS  
	This script shows Lync users last logon to a Lync pool based on the Lync CDR database and will display Lync Orphaned Users

.NOTES  
  Version      				: 0.4.5
  Rights Required			: Local admin
  Lync Version				: 2013 (Tested on February 2015 CU7 Update)
  Authors       			: Guy Bachar, Yoav Barzilay
  Last Update               : 25-May-2015
  Twitter/Blog	            : @GuyBachar, http://guybachar.us
  Twitter/Blog	            : @y0avb, http://y0av.me


.VERSION
  0.1 - Initial Version for connecting Internal Lync Servers
  0.2 - With the help of Anthony Caragol we were able to pull out more information using remoteing for powershell
  0.3 - Retrieving  Information on assigned certificates
  0.4 - Adding support for SFB and EDGE Servers, Fixing Connectivity issues and adding support for test-connection prior to the connectiviy
  0.4.5 - Fixing OWAS Certificates infromation from multiple servers
	
#>

param(
[Parameter(Position=0, Mandatory=$False) ][ValidateNotNullorEmpty()][switch] $EdgeCertificates,
[Parameter(Position=1, Mandatory=$False) ][ValidateNotNullorEmpty()][switch] $OWASCertificates,
[Parameter(Position=2, Mandatory=$False) ][ValidateNotNullorEmpty()][string] $FEPool
)

#region Script Information
Clear-Host
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Lync Certificates Reporter" -ForegroundColor Green
Write-Host "Version: 0.4.5" -ForegroundColor Green
Write-Host 
Write-Host "Authors:" -ForegroundColor Green
Write-Host " Guy Bachar       | @GuyBachar        | http://guybachar.us" -ForegroundColor Green
Write-Host " Yoav Barzilay    | @y0avb            | http://y0av.me" -ForegroundColor Green
Write-Host " Anthony Caragol  | @CAnthonyCaragol  | http://www.lyncfix.com" -ForegroundColor Green
Write-host
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
#endregion

#region Verifying Administrator Elevation
Write-Host Verifying User permissions... -ForegroundColor Yellow -NoNewline
#Start-Sleep -Seconds 2
#Verify if the Script is running under Admin privileges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
  [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
  Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!"
  Write-Host 
  Break
}
else
{
    Write-Host " Done!" -ForegroundColor Green
}
#endregion

#region Import Lync Module
Write-Host
Write-Host "Please wait while we're loading Lync PowerShell Module..." -ForegroundColor Yellow -NoNewline
  if(-not (Get-Module -Name "Lync")){
    if(Get-Module -Name "Lync" -ListAvailable){
      Import-Module -Name "Lync";
      #Write-Host "Loading Lync Module..." -ForegroundColor Yellow -NoNewline
    }
    else{
      Write-Host Lync Module does not exist on this computer, please verify the Lync Admin tools installed -ForegroundColor Red
      exit;   
    }    
  }

Write-Host -NoNewline    
Write-Host " Done!" -ForegroundColor Green
#Start-Sleep -Seconds 1
#endregion

# Retriving Pool and Servers Information
if ($FEPool.Length -eq 0) {$Poollist = Get-CsPool | Where-Object {(($_.Services -like "*Registrar*") -OR ($_.Services -like "*MediationServer*")) -and ($_.Site -ne "Site:BackCompatSite")}}
else {$Poollist = Get-CsPool | Where-Object {(($_.Services -like "*Registrar*") -OR ($_.Services -like "*MediationServer*")) -AND ($_.Identity -eq $FEPool)}}
$EDGElist = Get-CsPool | Where-Object {($_.Services -like "*EDGE*") -and ($_.Site -ne "Site:BackCompatSite")}
$OWASList = Get-CsPool | Where-Object {($_.Services -like "*WAC*")}
$FEServerList=$Poollist.computers
$EDGEServerList=$EDGElist.computers
$OWASServerList=$OWASList.computers
$ServerVersion = Get-CsServerVersion
$PSRemoteConnectionPort = "80"

#### Building HTML File ####
function CreateHtmFile
{
    $FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
    $ServicesFileName = $env:TEMP+"\LyncCertReport-"+$FileDate+".htm"
    $HTMLFile = New-Item -ItemType file $ServicesFileName -Force
    return $HTMLFile
}

Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Certificates Report</title>'
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
add-content $fileName  "<td colspan='7' height='25' align='center'>"
add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'><strong>Certificate Report</strong> - $date<BR>Server Version: $ServerVersion </font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr bgcolor=#0099CC>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Friendly Name / Usage</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Issuer</font></td>"
Add-Content $fileName "<td width='18%' align='center'><font color=#FFFFFF>Thumbprint</font></td>"
Add-Content $fileName "<td width='32%' align='center'><font color=#FFFFFF>Subject Name</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Issue Date</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Expiration Date</font></td>"
Add-Content $fileName "<td width='10%' align='center'><font color=#FFFFFF>Expires In</font></td>"
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
param($fileName,$FriendlyName,$Issuer,$Subject,$Thumbprint,$NotBefore,$NotAfter)
$TimeDiff = New-TimeSpan (Get-Date) $NotAfter
$DaysDiff = $TimeDiff.Days

 if ($NotAfter -gt (Get-date).AddDays(60))
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$FriendlyName</td>"
 Add-Content $fileName "<td>$Issuer</td>"
 Add-Content $fileName "<td>$Subject</td>"
 Add-Content $fileName "<td>$Thumbprint</td>"
 Add-Content $fileName "<td align='center'>$NotBefore</td>"
 Add-Content $fileName "<td align='center'>$NotAfter</td>"
 Add-Content $fileName "<td bgcolor='#00FF00' align=center>$DaysDiff days</td>"
 Add-Content $fileName "</tr>"
 }
 elseif ($NotAfter -lt (Get-date).AddDays(30))
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$FriendlyName</td>"
 Add-Content $fileName "<td>$Issuer</td>"
 Add-Content $fileName "<td>$Subject</td>"
 Add-Content $fileName "<td>$Thumbprint</td>"
 Add-Content $fileName "<td align='center'>$NotBefore</td>"
 Add-Content $fileName "<td align='center'>$NotAfter</td>"
 Add-Content $fileName "<td bgcolor='#FF0000' align=center>$DaysDiff days</td>"
 Add-Content $fileName "</tr>"
 }
 else
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$FriendlyName</td>"
 Add-Content $fileName "<td>$Issuer</td>"
 Add-Content $fileName "<td>$Subject</td>"
 Add-Content $fileName "<td>$Thumbprint</td>"
 Add-Content $fileName "<td align='center'>$NotBefore</td>"
 Add-Content $fileName "<td align='center'>$NotAfter</td>"
 Add-Content $fileName "<td bgcolor='#FBB917' align=center>$DaysDiff days</td>"
 Add-Content $fileName "</tr>"
 }
}

Function sendEmail
{ param($from,$to,$subject,$smtphost,$htmlFileName)
$body = Get-Content $htmlFileName
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
$msg.isBodyhtml = $true
$smtp.send($msg)
}

Function TestServersConnectivity
{ param($ServersArray)
    $TestedServers = @()
    foreach ($Computer in $ServersArray)
    {
        if (Test-Connection -ComputerName $Computer -Count 2 -Quiet)
        {
            Write-Host $Computer -ForegroundColor Green -NoNewline; Write-Host " - is accessible and will be tested for certificate expiration"
            $TestedServers += $Computer
        } 
        else
        {
            Write-Host $Computer -ForegroundColor Red -NoNewline; Write-Host " - is not accessible and will be removed from the list"
        }
    }
    return $TestedServers
}

# Main Script
$ServicesFileName = CreateHtmFile
writeHtmlHeader $ServicesFileName

#Front Ends Certificate list
try
{
    #Cleaning Servers that are not reachable   
    Write-Host "`nTesting Front End Connectivity:" -ForegroundColor Yellow
    $FETestedServerList = TestServersConnectivity ($FEServerList)

    Write-Host "`nRetrieving Certificate information from the Front End Servers..." -ForegroundColor Yellow
    $RemoteFECertList = Invoke-Command -ComputerName $FETestedServerList -ScriptBlock {Get-CsCertificate -ErrorAction SilentlyContinue}
}
catch
{
    Write-Host
    Write-Host "Error Conencting to local server $FETestedServerList, Please verify connectivity and permissions" -ForegroundColor Red
    Continue
}

# EDGE Certificates list
if ($EdgeCertificates.IsPresent)
{
    try
    {
    
    #EDGE Certificates List
    Write-Host "`nTesting EDGE Connectivity:" -ForegroundColor Yellow
    $EDGETestedServerList = TestServersConnectivity ($EDGEServerList)
    
    }
    catch
    {
        Write-Host
        Write-Host "Error Conencting to local server $EDGEServer, Please verify connectivity and permissions" -ForegroundColor Red
    Continue
    }

      
    #Validating EDGE Connectivity
    Write-Host "`nVerifying EDGE Servers are configured as Trusted Hosts..." -ForegroundColor Yellow
    $TrustedHosts = Get-Item WSMan:\localhost\Client\TrustedHosts
    if ($TrustedHosts.Value -eq "*")
    {
       Write-Host "`nThe following Trusted Hosts are configured:`n" -ForegroundColor Yellow -NoNewline; Write-Host $TrustedHosts.Value
    }
    else
    {
        Write-Host "`nConfiguring the following Trusted Hosts:" -ForegroundColor Yellow -NoNewline; Write-Host " *" -ForegroundColor Green
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
    }

    
    $RemoteEdgeCertList = @()
    Write-Host "`nRetrieving Certificate information from the EDGE Servers..." -ForegroundColor Yellow

    foreach ($EDGEServer in $EDGETestedServerList)
    {
        try
        {
        $S = New-PSSession -ComputerName $EDGEServer -Port $PSRemoteConnectionPort -Credential (Get-Credential -Message "Please enter your Edge Server's credentials" $EDGEServer\administrator) -ErrorAction Ignore
        $RemoteEDGEServerCertList = Invoke-Command -Session $S -Scriptblock {Get-CsCertificate} -ErrorAction Ignore
        Remove-PSSession $S
        }
        catch
        {
          Write-Host "Error Conencting to local server $EDGEServer, Please verify connectivity and permissions" -ForegroundColor Red  
        }

        $RemoteEdgeCertList += $RemoteEDGEServerCertList
    }
}


#Office Web AppsCertificate List
if ($OWASCertificates.IsPresent)
{
    try
    {
        #Cleaning Servers that are not reachable   
        Write-Host "`nTesting Office Web Apps Connectivity:" -ForegroundColor Yellow
        $OWASTestedServerList = TestServersConnectivity ($OWASServerList)

        Write-Host "`nRetrieving Certificate information from the Office Web Apps Servers..." -ForegroundColor Yellow
        
        foreach ($OWASServer in $OWASTestedServerList)
        {
            $Store = New-Object System.Security.Cryptography.X509Certificates.X509Store("$OWASServer\MY","LocalMachine") -ErrorAction Ignore
            $Store.Open("ReadOnly")
            $RemoteOWASCertList += $store.Certificates
        }

    }
    catch
    {
        Write-Host
        Write-Host "Error Conencting to local server $OWASServer, Please verify connectivity and permissions" -ForegroundColor Red
        #Continue
    }
}



#Setting Array for Uniqe Server Lists
$UniqeFEServersList = @()
$UniqeEDGEServersList = @()
$UniqeOWASServersList = $RemoteOWASCertList


foreach ($FEItem in $RemoteFECertList)
{
    $UniqeFEServersList += $FEItem.PSComputerName
}

foreach ($EdgeItem in $RemoteEdgeCertList)
{
    $UniqeEDGEServersList += $EdgeItem.PSComputerName
}

<#foreach ($OWASItem in $RemoteOWASCertList)
{
    $UniqeOWASServersList += $OWASItem.PSComputerName
}#>

$UniqeFEServersList = $UniqeFEServersList | Sort-Object -Unique
foreach ($Server in $UniqeFEServersList)
{       
        Add-Content $ServicesFileName "<table width='100%'><tbody>"
        Add-Content $ServicesFileName "<tr bgcolor='#0080FF'>"
        Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='segoe ui' color='#FFFFFF' size='2'><strong>$Server</strong></font></td>"
        Add-Content $ServicesFileName "</tr>"
        WriteTableHeader $ServicesFileName
        foreach ($item in $RemoteFECertList)
        {
	      if ($item.PSComputerName -eq $Server)
            {
                writeServiceInfo $ServicesFileName $item.Use $item.Issuer $item.Thumbprint $item.Subject $item.NotBefore $item.NotAfter
            }
        }
}

$UniqeEDGEServersList = $UniqeEDGEServersList | Sort-Object -Unique
foreach ($EDGEServer in $UniqeEDGEServersList)
{       
        Add-Content $ServicesFileName "<table width='100%'><tbody>"
        Add-Content $ServicesFileName "<tr bgcolor='#0080FF'>"
        Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='segoe ui' color='#FFFFFF' size='2'><strong>$EDGEServer</strong></font></td>"
        Add-Content $ServicesFileName "</tr>"
        WriteTableHeader $ServicesFileName
        foreach ($item in $RemoteEdgeCertList)
        {
	      if ($item.PSComputerName -eq $EDGEServer)
            {
                writeServiceInfo $ServicesFileName $item.Use $item.Issuer $item.Thumbprint $item.Subject $item.NotBefore $item.NotAfter
            }
        }
}

$UniqeOWASServersList = $UniqeOWASServersList | Sort-Object -Unique
foreach ($Server in $UniqeOWASServersList)
{
        $OWASServerName = $Server.subject
        Add-Content $ServicesFileName "<table width='100%'><tbody>"
        Add-Content $ServicesFileName "<tr bgcolor='#0080FF'>"
        Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='segoe ui' color='#FFFFFF' size='2'><strong>$OWASServerName (Office Web Apps Server)</strong></font></td>"
        Add-Content $ServicesFileName "</tr>"
        WriteTableHeader $ServicesFileName

        writeServiceInfo $ServicesFileName $server.FriendlyName $server.Issuer $server.Thumbprint $server.Subject $server.NotBefore $server.NotAfter
}


#Adding Twitter Account for Feedback
Add-Content $ServicesFileName "</table>"
Add-Content $ServicesFileName "<table width='100%'><tbody>"		
Add-Content $ServicesFileName "<tr bgcolor='#D8D8D8'>"		
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='segoe ui' color='#848484' size='2'>Script by <a href=https://twitter.com/guybachar target=_blank>Guy Bachar</a> and <a href=https://twitter.com/y0avb target=_blank>Yoav Barzilay</a></font></td>"		
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>"

writeHtmlFooter $ServicesFileName

### Configuring Email Parameters
#sendEmail from@domain.com to@domain.com "Services State Report - $Date" SMTPS_ERVER $ServicesFileName

#Closing HTML
writeHtmlFooter $ServicesFileName
Write-Host "`n`nThe File was generated at the following location: $ServicesFileName" -NoNewline; Write-Host "`nOpenning file..." -ForegroundColor Cyan
Invoke-Item $ServicesFileName