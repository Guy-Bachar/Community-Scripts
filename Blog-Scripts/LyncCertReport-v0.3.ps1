<#  
.SYNOPSIS  
	This script shows Lync users last logon to a Lync pool based on the Lync CDR database and will display Lync Orphaned Users

.NOTES  
  Version      				: 0.3
  Rights Required			: Local admin
  Lync Version				: 2013 (tested on August 2014 CU5 Update)
  Authors       			: Guy Bachar, Yoav Barzilay
  Last Update               : 8-August-2014
  Twitter/Blog	            : @GuyBachar, http://guybachar.us
  Twitter/Blog	            : @y0avb, http://y0av.me
  Twitter/Blog	            : @CAnthonyCaragol, http://www.lyncfix.com/



.VERSION
  0.1 - Initial Version for connecting Internal Lync Servers
  0.2 - With the help of Anthony Caragol we were able to pull out more information using remoteing for powershell
  0.3 - Pulling Information on assigned certificates
	
#>

#region Script Information
Clear-Host
Write-Host "--------------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Lync Certificates Reporter" -ForegroundColor Green
Write-Host "Version: 0.3" -ForegroundColor Green
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
Write-Host Verifying User permissions... -ForegroundColor Yellow
Start-Sleep -Seconds 2
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
Write-Host "Please wait while we're loading Lync PowerShell Module..." -ForegroundColor Yellow
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
Start-Sleep -Seconds 1
#endregion

$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LyncCertReport-"+$FileDate+".htm"
New-Item -ItemType file $ServicesFileName -Force
$Poollist = Get-CsPool | Where-Object {($_.Services -like "*Registrar*") -OR ($_.Services -like "*MediationServer*")}
$EDGElist = Get-CsPool | Where-Object {$_.Services -like "*EDGE*"}
$WAClist = Get-CsPool | Where-Object {($_.Services -like "*WAC*")}
$ServerList=$Poollist.computers
$EDGEServerList=$EDGElist.computers
$WacServerList=$WAClist.computers

#### Building HTML File ####
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
add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'>Lync Certificates Report - $date</font>"
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
 Add-Content $fileName "<td bgcolor='#00FF00' align=center>$DaysDiff</td>"
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
 Add-Content $fileName "<td bgcolor='#FF0000' align=center>$DaysDiff</td>"
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
 Add-Content $fileName "<td bgcolor='#FBB917' align=center>$DaysDiff</td>"
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

# Main Script
writeHtmlHeader $ServicesFileName

try
{
$RemoteCertList = Invoke-Command -ComputerName $serverlist -ScriptBlock {Get-CsCertificate -ErrorAction SilentlyContinue}
#$RemoteEdgeCertList = Invoke-Command -ComputerName $EDGEServerList -ScriptBlock {Get-CsCertificate} -Authentication Credssp -Credential (Get-Credential -Credential $EDGEServerList\administrator)
}
catch
{
    Write-Host
    Write-Host "Error Conencting to local server $FQDN, Please verify connectivity and permissions" -ForegroundColor Red
    Continue
}

$UniqeServersList = @()
foreach ($item in $RemoteCertList)
{
    $UniqeServersList += $item.PSComputerName
}
$UniqeServersList = $UniqeServersList | Sort-Object -Unique

foreach ($Server in $UniqeServersList)
{       
        Add-Content $ServicesFileName "<table width='100%'><tbody>"
        Add-Content $ServicesFileName "<tr bgcolor='#0099CC'>"
        Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='segoe ui' color='#FFFFFF' size='2'>$Server</font></td>"
        Add-Content $ServicesFileName "</tr>"
        WriteTableHeader $ServicesFileName
        foreach ($item in $RemoteCertList)
        {
	      if ($item.PSComputerName -eq $Server)
            {
                writeServiceInfo $ServicesFileName $item.Use $item.Issuer $item.Thumbprint $item.Subject $item.NotBefore $item.NotAfter
            }
        }
        Add-Content $ServicesFileName "</table>"
}

foreach ($wacserver in $WacServerList)
{
    try
    {
        $Store = New-Object System.Security.Cryptography.X509Certificates.X509Store("$wacserver\MY","LocalMachine") -ErrorAction Stop
        $Store.Open("ReadOnly")
        $Certificates = $store.Certificates
    }
    catch
    {
        Write-Host
        Write-Host "Error Conencting to local server $FQDN, Please verify connectivity and permissions" -ForegroundColor Red
        Continue
    }
    Add-Content $ServicesFileName "<table width='100%'><tbody>"
    Add-Content $ServicesFileName "<tr bgcolor='#0099CC'>"
    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='segoe ui' color='#FFFFFF' size='2'>$wacserver</font></td>"
    Add-Content $ServicesFileName "</tr>"
    WriteTableHeader $ServicesFileName

    foreach ($item in $Certificates)
    {
	    writeServiceInfo $ServicesFileName $item.FriendlyName $item.Issuer $item.Thumbprint $item.Subject $item.NotBefore $item.NotAfter
    }
    Add-Content $ServicesFileName "</table>"
}

writeHtmlFooter $ServicesFileName

### Configuring Email Parameters
#sendEmail from@domain.com to@domain.com "Services State Report - $Date" SMTPS_ERVER $ServicesFileName

#Closing HTML
writeHtmlFooter $ServicesFileName
Write-Host "`n`nThe File was generated at the following location: $ServicesFileName `n`nOpenning file..." -ForegroundColor Cyan
Invoke-Item $ServicesFileName