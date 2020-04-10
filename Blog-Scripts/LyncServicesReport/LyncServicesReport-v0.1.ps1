<#
.SYNOPSIS
    The script create an HTML based reports of all Lync related services based on an input text file contains servers list.

.DESCRIPTION
    	
	This script will generate an htm file which will include a detailed report of Lync services status.
	The script will only display Lync releated services (Lync and SQL) and their current state.
    Services which were orginally configured to set for "Automatic" state and currently are in "Stopped" state will be marked in Red.
    The script also support sending the HTM file as an email body, by setting the parameters in the Email section.
	
.NOTES
    File Name: LyncServicesReport.ps1
	Version: 0.1
	Last Update: 17-Apr-2014
    Author: Guy Bachar, http://guybachar.us
    The script are provided “AS IS” with no guarantees, no warranties, USE ON YOUR OWN RISK.    
#>


$ServicesFileName = "C:\Scripts\LyncServicesReport.htm"
$serverlist = "C:\Script\Servers.txt"
New-Item -ItemType file $ServicesFileName -Force

Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Servers Services Report</title>'
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
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Lync Servers Services Report - $date</strong></font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr bgcolor=#CCCCCC>"
Add-Content $fileName "<td width='10%' align='center'>Server Name</td>"
Add-Content $fileName "<td width='40%' align='center'>Services Name</td>"
Add-Content $fileName "<td width='10%' align='center'>State</td>"
Add-Content $fileName "<td width='20%' align='center'>User Name</td>"
Add-Content $fileName "<td width='10%' align='center'>Status</td>"
Add-Content $fileName "<td width='10%' align='center'>Start Mode</td>"
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
param($fileName,$SystemName,$DisplayName,$StartMode,$State,$StartName,$Status,$Description)
 if ($State -eq "Running")
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$SystemName</td>"
 Add-Content $fileName "<td>$DisplayName</td>"
 Add-Content $fileName "<td bgcolor='#00FF00' align=center>$State</td>"
 Add-Content $fileName "<td>$StartName</td>"
 Add-Content $fileName "<td>$Status</td>"
 Add-Content $fileName "<td>$StartMode</td>"
 Add-Content $fileName "</tr>"
 }
 elseif ($State -eq "Stopped")
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$SystemName</td>"
 Add-Content $fileName "<td>$DisplayName</td>"
 Add-Content $fileName "<td bgcolor='#FF0000' align=center>$State</td>"
 Add-Content $fileName "<td>$StartName</td>"
 Add-Content $fileName "<td>$Status</td>"
 Add-Content $fileName "<td>$StartMode</td>"
 Add-Content $fileName "</tr>"
 }
 else
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$SystemName</td>"
 Add-Content $fileName "<td>$DisplayName</td>"
 Add-Content $fileName "<td bgcolor='#FBB917' align=center>$State</td>"
 Add-Content $fileName "<td>$StartName</td>"
 Add-Content $fileName "<td>$Status</td>"
 Add-Content $fileName "<td>$StartMode</td>"
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

writeHtmlHeader $ServicesFileName
foreach ($server in Get-Content $serverlist)
{
 Add-Content $ServicesFileName "<table width='100%'><tbody>"
 Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
 Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> $server </strong></font></td>"
 Add-Content $ServicesFileName "</tr>"

 writeTableHeader $ServicesFileName

 $ServicesList = Get-WmiObject -Class Win32_Service -ComputerName $server | Where-Object { (($_.DisplayName -like'Lync*') -or ($_.DisplayName -like'SQL*')) -and ($_.StartMode -eq "Auto") }
 
 foreach ($item in $ServicesList)
 {
  writeServiceInfo $ServicesFileName $item.SystemName $item.DisplayName $item.StartMode $item.State $item.StartName $item.Status $item.Description
 }
Add-Content $ServicesFileName "</table>"
}
writeHtmlFooter $ServicesFileName
$date = ( get-date ).ToString('yyyy/MM/dd')

### Configuring Email Parameters
sendEmail from@domain.com to@domain.com "Services State Report - $Date" SMTPS_ERVER $ServicesFileName
