<#
.SYNOPSIS
  The script will output all Lync Releated QoS Settings Configured.

.DESCRIPTION
  	
	
	
.NOTES
  File Name: Get-CsQoSReport.ps1
	Version: 0.1
	Last Update: 25-May-2014
  Author: Guy Bachar, @GuyBachar, http://guybachar.us"
  Author: Yoav Barzilay, @y0av, http://y0av.me/"
  The script are provided “AS IS” with no guarantees, no warranties, USE ON YOUR OWN RISK.  

.WHATSNEW
  0.1 - HTML QoS Report Added
#> 

Clear-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Get Lync QoS Configurations" -ForegroundColor Green
Write-Host "Version: 0.1" -ForegroundColor Green
Write-Host 
Write-Host "Authors:" -ForegroundColor Green
Write-Host " Guy Bachar    | @GuyBachar | http://guybachar.us" -ForegroundColor Green
Write-Host " Yoav Barzilay | @y0avb     | http://y0av.me" -ForegroundColor Green
Write-host
$Date = Get-Date -DisplayHint DateTime
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Data collected:" , $Date -ForegroundColor Yellow
Write-host


#Variables
$QoSConferencingServer = Get-CsService -ConferencingServer | Select-Object Identity, AudioPortStart, AudioPortCount, VideoPortStart, VideoPortCount, AppSharingPortStart, AppSharingPortCount
$QoSApplicationServer = Get-CsService -ApplicationServer | Select-Object Identity, AudioPortStart, AudioPortCount
$QoSMediationServer = Get-CsService -MediationServer | Select-Object Identity, AudioPortStart, AudioPortCount
$QoSMediaConfiguration = Get-CsMediaConfiguration
$QoSUCPhoneConfiguration = Get-CsUCPhoneConfiguration
$QoSCsConferencingConfiguration = Get-CsConferencingConfiguration


$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LyncQoSReport-"+$FileDate+".htm"

#Import Lync Module
  if(-not (Get-Module -Name "Lync")){
    if(Get-Module -Name "Lync" -ListAvailable){
      Import-Module -Name "Lync";
      Write-Host "Loading Lync Module";
    }
    else{
      Write-Host "Lync Module does not exist on this computer, please verify the Lync Admin tools installed";    
      exit;   
    }    
  }

#Verify if the Script is running under Admin privliges
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
  [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
  Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
  Write-Host 
  Break
}




#### Building HTML File ####
Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Quality Of Service (QOS) Report</title>'
add-content $fileName '<STYLE TYPE="text/css">'
add-content $fileName  "<!--"
add-content $fileName  "td {"
add-content $fileName  "font-family: Segoe UI;"
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
add-content $fileName  "<tr bgcolor='#336699 '>"
add-content $fileName  "<td colspan='7' height='25' align='center'>"
add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'><strong>Lync QoS Report - $date</strong></font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function WriteQoSConferencingServer
{
param($fileName,$Identity, $AudioPortStart, $AudioPortCount, $VideoPortStart, $VideoPortCount, $AppSharingPortStart, $AppSharingPortCount)
    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='28%' align='center'>Server Name</td>"
    Add-Content $fileName "<td width='12%' align='center'>AudioPortStart</td>"
    Add-Content $fileName "<td width='12%' align='center'>AudioPortCount</td>"
    Add-Content $fileName "<td width='12%' align='center'>VideoPortStart</td>"
    Add-Content $fileName "<td width='12%' align='center'>VideoPortCount</td>"
    Add-Content $fileName "<td width='12%' align='center'>AppSharingPortStart</td>"
    Add-Content $fileName "<td width='12%' align='center'>AppSharingPortCount</td>"
    Add-Content $fileName "</tr>"
    Add-Content $fileName "<tr>"
    Add-Content $fileName "<td>$Identity</td>"
    Add-Content $fileName "<td align=center>$AudioPortStart</td>"
    Add-Content $fileName "<td align=center>$AudioPortCount</td>"
    Add-Content $fileName "<td align=center>$VideoPortStart</td>"
    Add-Content $fileName "<td align=center>$VideoPortCount</td>"
    Add-Content $fileName "<td align=center>$AppSharingPortStart</td>"
    Add-Content $fileName "<td align=center>$AppSharingPortCount</td>"
    Add-Content $fileName "</tr>"
}

Function WriteQoSApplicationServer
{
param($fileName,$Identity, $AudioPortStart, $AudioPortCount)
    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='28%' align='center'>Server Name</td>"
    Add-Content $fileName "<td width='36%' align='center'>AudioPortStart</td>"
    Add-Content $fileName "<td width='36%' align='center'>AudioPortCount</td>"
    Add-Content $fileName "</tr>"
    Add-Content $fileName "<tr>"
    Add-Content $fileName "<td>$Identity</td>"
    Add-Content $fileName "<td align=center>$AudioPortStart</td>"
    Add-Content $fileName "<td align=center>$AudioPortCount</td>"
    Add-Content $fileName "</tr>"
}

Function WriteQoSMediationServer
{
param($fileName,$Identity, $AudioPortStart, $AudioPortCount)
    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='28%' align='center'>Server Name</td>"
    Add-Content $fileName "<td width='36%' align='center'>AudioPortStart</td>"
    Add-Content $fileName "<td width='36%' align='center'>AudioPortCount</td>"
    Add-Content $fileName "</tr>"
    Add-Content $fileName "<tr>"
    Add-Content $fileName "<td>$Identity</td>"
    Add-Content $fileName "<td align=center>$AudioPortStart</td>"
    Add-Content $fileName "<td align=center>$AudioPortCount</td>"
    Add-Content $fileName "</tr>"
}

Function WriteQoSMediaConfiguration_QoSUCPhoneConfiguration
{
param($fileName,$Identity, $EnableQoS, $VoiceDiffServTag, $Voice8021p)
    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='25%' align='center'>Policy Name</td>"
    Add-Content $fileName "<td width='25%' align='center'>QoS Enabled for Lync?</td>"
    Add-Content $fileName "<td width='25%' align='center'>UC Phone DSCP Value</td>"
    Add-Content $fileName "<td width='25%' align='center'>Voice 802.1p Value</td>"
    Add-Content $fileName "</tr>"
    Add-Content $fileName "<tr>"
    Add-Content $fileName "<td align=center>$Identity</td>"
    if ($EnableQoS -eq "True")
    { Add-Content $fileName "<td align=center bgcolor='#33CC33'>$EnableQoS</td>" }
    else { Add-Content $fileName "<td align=center bgcolor='#FF0000'>$EnableQoS</td>" }
    Add-Content $fileName "<td align=center>$VoiceDiffServTag</td>"
    Add-Content $fileName "<td align=center>$Voice8021p</td>"
    Add-Content $fileName "</tr>"
}

Function WriteQoSCsConferencingConfiguration
{
param($fileName,$Identity, $ClientMediaPortRangeEnabled, $ClientMediaPort ,$ClientMediaPortRange ,$ClientAudioPort ,$ClientAudioPortRange ,$ClientVideoPort ,$ClientVideoPortRange, $ClientAppSharingPort, $ClientAppSharingPortRange)
    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='10%' align='center'>Policy Name</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientMediaPortRangeEnabled</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientMediaPort</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientMediaPortRange</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientAudioPort</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientAudioPortRange</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientVideoPort</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientVideoPortRange</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientAppSharingPort</td>"
    Add-Content $fileName "<td width='10%' align='center'>ClientAppSharingPortRange</td>"
    Add-Content $fileName "</tr>"
    Add-Content $fileName "<tr>"
    Add-Content $fileName "<td>$Identity</td>"
    Add-Content $fileName "<td align=center>$ClientMediaPortRangeEnabled</td>"
    Add-Content $fileName "<td align=center>$ClientMediaPort</td>"
    Add-Content $fileName "<td align=center>$ClientMediaPortRange</td>"
    Add-Content $fileName "<td align=center>$ClientAudioPort</td>"
    Add-Content $fileName "<td align=center>$ClientAudioPortRange</td>"
    Add-Content $fileName "<td align=center>$ClientVideoPort</td>"
    Add-Content $fileName "<td align=center>$ClientVideoPortRange</td>"
    Add-Content $fileName "<td align=center>$ClientAppSharingPort</td>"
    Add-Content $fileName "<td align=center>$ClientAppSharingPortRange</td>"
    Add-Content $fileName "</tr>"
}


Function writeHtmlFooter
{
param($fileName)
Add-Content $fileName "</body>"
Add-Content $fileName "</html>"
}

#Open HTML
writeHtmlHeader $ServicesFileName

#Global Configurations
Add-Content $ServicesFileName "<table width='100%'><tbody>"
Add-Content $ServicesFileName "<tr bgcolor='#000099'>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=2><font face='Segoe UI' color='#FFFFFF' size='3'><strong> Servers Configurations </strong></font></td>"
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>"

#1. QoSConferencingServer
Add-Content $ServicesFileName "<table width='100%'><tbody>"
Add-Content $ServicesFileName "<tr bgcolor='#0099FF'>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=7><font face='Segoe UI' color='#FFFFFF' size='2'><strong> QoS Conferencing Server Policy </strong></font></td>"
foreach ($object in $QoSConferencingServer)
  {
  WriteQoSConferencingServer $ServicesFileName $object.Identity $object.AudioPortStart $object.AudioPortCount $object.VideoPortStart $object.VideoPortCount $object.AppSharingPortStart $object.AppSharingPortCount
  }
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>" 


#2. QoSApplicationServer
Add-Content $ServicesFileName "<table width='100%'>"
Add-Content $ServicesFileName "<tr bgcolor='#0099FF'>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=3><font face='Segoe UI' color='#FFFFFF' size='2'><strong> Application Server Policy </strong></font></td>"
foreach ($object in $QoSApplicationServer)
  {
  WriteQoSApplicationServer $ServicesFileName $object.Identity $object.AudioPortStart $object.AudioPortCount
  }
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>" 

#3. QoSMediationServer
Add-Content $ServicesFileName "<table width='100%'>"
Add-Content $ServicesFileName "<tr bgcolor='#0099FF'>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=3><font face='Segoe UI' color='#FFFFFF' size='2'><strong> Mediation Server Policy </strong></font></td>"
foreach ($object in $QoSMediationServer)
  {
  WriteQoSMediationServer $ServicesFileName $object.Identity $object.AudioPortStart $object.AudioPortCount
  }
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>" 

#Global Configurations
Add-Content $ServicesFileName "<table width='100%'><tbody>"
Add-Content $ServicesFileName "<tr bgcolor='#000099'><BR>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=2><font face='Segoe UI' color='#FFFFFF' size='3'><strong> Global Configurations </strong></font></td>"
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>" 

#4. QoSMediaConfiguration & QoSUCPhoneConfiguration
Add-Content $ServicesFileName "<table width='100%'><tbody>"
Add-Content $ServicesFileName "<tr bgcolor='#0099FF'>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=4><font face='Segoe UI' color='#FFFFFF' size='2'><strong> Media Configuration & UC Phone Configuration Policy </strong></font></td>"
WriteQoSMediaConfiguration_QoSUCPhoneConfiguration $ServicesFileName $QoSMediaConfiguration.Identity $QoSMediaConfiguration.EnableQoS $QoSUCPhoneConfiguration.VoiceDiffServTag $QoSUCPhoneConfiguration.Voice8021p
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>"

#5. QoSCsConferencingConfiguration
Add-Content $ServicesFileName "<table width='100%'><tbody>"
Add-Content $ServicesFileName "<tr bgcolor='#0099FF'>"
Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=10><font face='Segoe UI' color='#FFFFFF' size='2'><strong> Conferencing Configuration Policy </strong></font></td>"
WriteQoSCsConferencingConfiguration $ServicesFileName $QoSCsConferencingConfiguration.Identity $QoSCsConferencingConfiguration.ClientMediaPortRangeEnabled $QoSCsConferencingConfiguration.ClientMediaPort `
$QoSCsConferencingConfiguration.ClientMediaPortRange $QoSCsConferencingConfiguration.ClientAudioPort $QoSCsConferencingConfiguration.ClientAudioPortRange $QoSCsConferencingConfiguration.ClientVideoPort `
$QoSCsConferencingConfiguration.ClientVideoPortRange $QoSCsConferencingConfiguration.ClientAppSharingPort $QoSCsConferencingConfiguration.ClientAppSharingPortRange
Add-Content $ServicesFileName "</tr>"
Add-Content $ServicesFileName "</table>" 


#Closing HTML
writeHtmlFooter $ServicesFileName
Write-Host "The File was generated at the following location: $ServicesFileName `n`nOpenning file..." -ForegroundColor Cyan
Invoke-Item $ServicesFileName
