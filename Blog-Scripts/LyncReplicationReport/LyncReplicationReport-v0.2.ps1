<#
.SYNOPSIS
   	The script create an HTML based reports of all Lync Replication Status

.DESCRIPTION
    	
	This script will generate an htm file which will include a detailed report of Lync replication status.
   	The script also support sending the HTM file as an email body, by setting the parameters in the Email section.
	
.NOTES
   	File Name: LyncReplicationReport.ps1
	Version: 0.2
   	Author: Guy Bachar, @GuyBachar, "http://guybachar.us/"
   	Author: Yoav Barzilay, @y0av, "http://y0av.me/"
   	The script are provided “AS IS” with no guarantees, no warranties, USE ON YOUR OWN RISK. 

.VERSION
   	v0.1 - 23-April-2014 - Created for Lync 2010 & Lync 2013 environments
   	v0.2 - 08-May-2014 	 - Added options selection, added module verifications
	
.SOURCES
	- Format Color Functions - http://bgreco.net/powershell/format-color/
	- Source HTML File - http://gallery.technet.microsoft.com/scriptcenter/6e935887-6b30-4654-b977-6f5d289f3a63
     
#>


[CmdletBinding(SupportsShouldProcess = $True)]
param (
	# Defines the From email Address.
	[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
	[ValidateNotNullOrEmpty()]
	[string] $MailFrom,
	
	# Defines the to email Address
	[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
	[ValidateNotNullOrEmpty()]
	[string] $MailTo,
	
	# Defines the SMTP server Address
	[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
	[ValidateNotNullOrEmpty()]
	[string] $SMTPServerName
)

#Script Info
Clear-Host
Write-Host "-------------------------------------------------------"
Write-Host
Write-Host "Lync Replication Report"
Write-Host
Write-Host "Version: 0.2"
Write-Host 
Write-Host "Authors:"
Write-Host
Write-Host " Guy Bachar        @GuyBachar     http://guybachar.us"
Write-Host " Yoav Barzilay     @y0avb         http://y0av.me"
Write-host
Write-Host "-------------------------------------------------------"
Write-Host
Write-Host

if (($MailFrom.Length -gt 0) -AND ($MailTo.Length -gt 0) -AND ($SMTPServerName.Length -gt 0))
{
$UserInput = 6
}

else
{
# Setting Parameters for Display or Export
Write-Host "****************************************************"
Write-Host "Output Selection (Default is output to screen)" -ForegroundColor DarkCyan
Write-Host
Write-Host "1)     Send an Email"
Write-Host "2)     Export to GridView"
Write-Host "3)     Export to CSV"
Write-Host "4)     Export to HTML"
Write-Host "5)     Export to Screen"
Write-Host
$UserInput = Read-Host "Please Enter your choice"
Write-Host
switch ($UserInput) 
    { 
        1 {"You chose: Export to Email"} 
        2 {"You chose: Export to GridView"} 
        3 {"You chose: Export to CSV"}
        4 {"You chose: Export to HTML"}
	    5 {"You chose: Export to Screen"}
	   default {"No option selected. Results will be displayed on the screen."}
    }

}

$ReplicationStatus = Get-CsManagementStoreReplicationStatus
if ($ReplicationStatus.Count -eq 0)		{
		Write-Host "`nerror: Problem finding replications`n`n" -ForegroundColor red
		break
	}


$date = ( get-date ).ToString('yyyy/MM/dd')
$filedate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LyncReplicationReport-"+$filedate+".htm"

Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Replication Report</title>'
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
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Lync Replication Report - $date</strong></font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr bgcolor=#CCCCCC>"
Add-Content $fileName "<td width='20%' align='center'>Server Name</td>"
Add-Content $fileName "<td width='20%' align='center'>Up to date</td>"
Add-Content $fileName "<td width='20%' align='center'>Last Status Report</td>"
Add-Content $fileName "<td width='20%' align='center'>Last Update Creation</td>"
Add-Content $fileName "<td width='20%' align='center'>Production Version</td>"
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
param($fileName,$ReplicaFQDN,$UpToDate,$LastStatusReport,$LastUpdateCreation,$ProductVersion)
 if ($UpToDate -eq $true)
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$ReplicaFQDN</td>"
 Add-Content $fileName "<td bgcolor='#00FF00' align=center>$UpToDate</td>"
 Add-Content $fileName "<td align=center>$LastUpdateCreation</td>"
 Add-Content $fileName "<td align=center>$LastStatusReport</td>"
 Add-Content $fileName "<td align=center>$ProductVersion</td>"
 Add-Content $fileName "</tr>"
 }
 elseif ($UpToDate -eq $false)
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$ReplicaFQDN</td>"
 Add-Content $fileName "<td bgcolor='#FF0000' align=center>$UpToDate</td>"
 Add-Content $fileName "<td align=center>$LastUpdateCreation</td>"
 Add-Content $fileName "<td align=center>$LastStatusReport</td>"
 Add-Content $fileName "<td align=center>$ProductVersion</td>"
 Add-Content $fileName "</tr>"
 }
 else
 {
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$ReplicaFQDN</td>"
 Add-Content $fileName "<td bgcolor='#FBB917' align=center>$UpToDate</td>"
 Add-Content $fileName "<td align=center>$LastUpdateCreation</td>"
 Add-Content $fileName "<td align=center>$LastStatusReport</td>"
 Add-Content $fileName "<td align=center>$ProductVersion</td>"
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


function Format-Color([hashtable] $Colors = @{}, [switch] $SimpleMatch) {
	$lines = ($input | Out-String) -replace "`r", "" -split "`n"
	foreach($line in $lines) {
		$color = ''
		foreach($pattern in $Colors.Keys){
			if(!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
			elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
		}
		if($color) {
			Write-Host -ForegroundColor $color $line
		} else {
			Write-Host $line
		}
	}
}

#Main Script
writeHtmlHeader $ServicesFileName
Add-Content $ServicesFileName "<table width='100%'><tbody>"
writeTableHeader $ServicesFileName           
        foreach ($item in $ReplicationStatus)
        {
            writeServiceInfo $ServicesFileName $item.ReplicaFQDN $item.UpToDate $item.LastStatusReport $item.LastUpdateCreation $item.ProductVersion
        }
	Add-Content $ServicesFileName "</table>"
 
writeHtmlFooter $ServicesFileName

If ($UserInput -eq 1)
	{
		$MailFrom = Read-Host "Please Enter FROM e-mail address"
		$MailTo 	= Read-Host "Please Enter TO e-mail address"
		$SMTPServerName = Read-Host "Please Enter SMTP Server Name or IP"
		sendEmail $MailFrom $MailTo "Lync Replication State Report - $Date" $SMTPServerName $ServicesFileName
     }
elseif ($UserInput -eq 2)     
	{
		$ReplicationStatus | Out-GridView
	}
elseif ($UserInput -eq 3)     
	{
		$CSVFileName = $env:TEMP+"\LastLogonExport-"+$filedate+".csv"
		$ReplicationStatus | Export-Csv -Path $CSVFileName
		Write-Host "The File is located under $CSVFileName"
	}
elseif ($UserInput -eq 4)     
	{
		Write-Host "The File is located under $ServicesFileName"
        Invoke-Item $ServicesFileName
	}
elseif ($UserInput -eq 6)     
	{
		sendEmail $MailFrom $MailTo "Lync Replication State Report - $Date" $SMTPServerName $ServicesFileName
	}
else
    { 
       Get-CsManagementStoreReplicationStatus | Format-Table -AutoSize | Format-Color @{'False' = 'Red'; 'True' = 'Green';'UpToDate' = 'Yellow'}
    }
