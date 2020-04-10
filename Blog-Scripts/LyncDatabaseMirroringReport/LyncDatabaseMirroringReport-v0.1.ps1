
param(
[Parameter(Mandatory=$False)][ValidateNotNullorEmpty()][switch] $SendEmail
)

import-module lync

Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Database Mirroring Report</title>'
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
add-content $fileName  "<font face='tahoma' color='#003399' size='4'><strong>Lync Database Mirroring Report - $date</strong></font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr bgcolor=#CCCCCC>"
Add-Content $fileName "<td width='10%' align='center'>Application</td>"
Add-Content $fileName "<td width='20%' align='center'>Database</td>"
Add-Content $fileName "<td width='20%' align='center'>Pool</td>"
Add-Content $fileName "<td width='20%' align='center'>Primary State</td>"
Add-Content $fileName "<td width='20%' align='center'>Primary DB Server</td>"
Add-Content $fileName "<td width='20%' align='center'>Mirror DB Server</td>"
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
param($fileName,$Application,$Database,$Pool,$PrimaryState,$PrimaryDB, $MirrorDB)
 
 Add-Content $fileName "<tr>"
 Add-Content $fileName "<td>$Application</td>"
 Add-Content $fileName "<td align=center>$Database</td>"
 Add-Content $fileName "<td align=center>$Pool</td>"
 if ($PrimaryState -eq "Principal")
    { 
        Add-Content $fileName "<td bgcolor='#00FF00' align=center>$PrimaryState</td>"
    }
else 
    {
        Add-Content $fileName "<td align=center>$PrimaryState</td>"
    }
 Add-Content $fileName "<td align=center>$PrimaryDB</td>"
 Add-Content $fileName "<td align=center>$MirrorDB</td>"
 Add-Content $fileName "</tr>"
 }


Function Get-LyncDatabaseMirror
{

Begin

{

$pools = @()

$dbServices = @{

'rgsconfig' = 'Application'

'rgsdyn' = 'Application'

'cpsdyn' = 'Application'

'lcslog' = 'Archiving'

'xds' = 'CentralMgmt'

'lis' = 'CentralMgmt'

'mgc' = 'PersistentChat'

'mgccomp' = 'PersistentChatCompliance'

'rtcab' = 'UserServer'

'rtcxds' = 'UserServer'

'rtcshared' = 'UserServer'

'lcscdr' = 'Monitoring'

'qoemetrics' = 'Monitoring'

}

$dbs = @()

$pools = @()

}

Process

{}

End

{

$pools = ((Get-CsTopology).Clusters | 

 Where {($_.RequiresReplication) -and (!$_.IsOnEdge)} | 

 %{[string]$_.Fqdn})

Foreach ($pool in $pools)

{

$dbstates = Get-CsDatabaseMirrorState -PoolFqdn $pool

if ($dbstates -ne $null)

{

foreach ($dbstate in $dbstates)

{

switch ($dbServices[[string]$dbstate.DatabaseName]) {

'Application' {

Get-CsService -PoolFqdn $pool -ApplicationServer | %{

$PrimaryDBServer = (([string]$_.ApplicationDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorApplicationDatabase).Split(':'))[1]

}

}

'Archiving' {

Get-CsService -PoolFqdn $pool -Registrar | %{

$PrimaryDBServer = (([string]$_.ArchivingDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorArchivingDatabase).Split(':'))[1]

}

}

'Monitoring' {

Get-CsService -PoolFqdn $pool -Registrar | %{

$PrimaryDBServer = (([string]$_.MonitoringDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorMonitoringDatabase).Split(':'))[1]

}

}

'CentralMgmt' {

Get-CsService -PoolFqdn $pool -CentralManagement | %{

$PrimaryDBServer = (([string]$_.CentralManagementDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorCentralManagementDatabase).Split(':'))[1]

}

}

'PersistentChat' {

Get-CsService -PoolFqdn $pool -PersistentChatServer | %{

$PrimaryDBServer = (([string]$_.PersistentChatDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorPersistentChatDatabase).Split(':'))[1]

}

}

'PersistentChatCompliance' {

Get-CsService -PoolFqdn $pool -PersistentChatServer | %{

$PrimaryDBServer = (([string]$_.PersistentChatComplianceDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorPersistentChatComplianceDatabase).Split(':'))[1]

}

}

'UserServer' {

Get-CsService -PoolFqdn $pool -UserServer | %{

$PrimaryDBServer = (([string]$_.UserDatabase).Split(':'))[1]

$MirrorDBServer = (([string]$_.MirrorUserDatabase).Split(':'))[1]

}

}

default {

$PrimaryDBServer = ''

$MirrorDBServer = ''

}

}

$dbprops = @{

'Pool' = $pool

'Application' = $dbServices[[string]$dbstate.DatabaseName]

'Database' = $dbstate.DatabaseName

'PrimaryState' = $dbstate.StateOnPrimary

'MirrorState' = $dbstate.StateOnMirror

'PrimaryDBServer' = $PrimaryDBServer

'MirrorDBServer' = $MirrorDBServer

}

New-Object psobject -Property $dbprops

}

}

}

}

}


function New-AuditLogReport {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $AuditLogEntry	
        )
	begin {
$css = @'
	<style type="text/css">
	body { font-family: Tahoma, Geneva, Verdana, sans-serif;}
	table {border-collapse: separate; border: 3px solid #103E69; caption-side: bottom;}
	td { border:1px solid #103E69; margin: 3px; padding: 3px; vertical-align: top; font-size: 12px;}
	thead th {background: #903; color:#fefdcf; text-align: left; font-weight: bold; padding: 3px;border: 1px solid #990033;}
	th {border:1px solid #CC9933; padding: 3px;}
	tbody th:hover {background-color: #fefdcf;}
	th a:link, th a:visited {color:#903; font-weight: normal; text-decoration: none; border-bottom:1px dotted #c93;}
	caption {background: #903; color:#fcee9e; padding: 4px 0; text-align: center; width: 40%; font-weight: bold;}
	tbody td a:link {color: #903;}
	tbody td a:visited {color:#633;}
	tbody td a:hover {color:#000; text-decoration: none;
	}
	</style>
'@	
		$sb = New-Object System.Text.StringBuilder
		[void]$sb.AppendLine($css)
		[void]$sb.AppendLine("<table cellspacing='0'>")
		[void]$sb.AppendLine("<tr><td align=center colspan='6'><strong>Lync Database Mirroring Report for $((get-date).ToShortDateString())</strong></td></tr>")
		[void]$sb.AppendLine("<tr>")
		[void]$sb.AppendLine("<td><strong>Application</strong></td>")
		[void]$sb.AppendLine("<td><strong>Database</strong></td>")
		[void]$sb.AppendLine("<td><strong>Pool</strong></td>")
		[void]$sb.AppendLine("<td><strong>Primary State</strong></td>")
		[void]$sb.AppendLine("<td><strong>Primary DB Server</strong></td>")
		[void]$sb.AppendLine("<td><strong>Mirror DB Server</strong></td>")
		[void]$sb.AppendLine("</tr>")
	}
	
	process {
		[void]$sb.AppendLine("<tr>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.Application.split("/")[-1])</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.Database.ToString())</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.Pool)</td>")
        if ($AuditLogEntry.PrimaryState -eq  "Principal")
        {
		    [void]$sb.AppendLine("<td><font color=00CC00>$($AuditLogEntry.PrimaryState)</font></td>")
        }
        else
        {
            [void]$sb.AppendLine("<td><font color=FF0000>$($AuditLogEntry.PrimaryState)</font></td>")
        }
        [void]$sb.AppendLine("<td>$($AuditLogEntry.PrimaryDBServer)</td>")
		[void]$sb.AppendLine("<td>$($AuditLogEntry.MirrorDBServer)</td>")
		[void]$sb.AppendLine("</tr>")
		$cmdletparameters = $null
	}
	
	end {
		[void]$sb.AppendLine("</table>")
		Write-Output $sb.ToString()
	}
}
 


$DatabaseMirror=Get-LyncDatabaseMirror
$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
#$Report = Get-LyncDatabaseMirror | New-AuditLogReport
#$Report | Out-File $env:TEMP"\DatabaseMirrorReport-"$FileDate".html"


$ReplicationStatus = Get-CsManagementStoreReplicationStatus
$date = ( get-date ).ToString('yyyy/MM/dd')
$filedate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LyncDatabaseMirrorReport-"+$filedate+".htm"

writeHtmlHeader $ServicesFileName

    Add-Content $ServicesFileName "<table width='100%'><tbody>"
    Add-Content $ServicesFileName "<tr bgcolor='#CCCCCC'>"
    Add-Content $ServicesFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> DatabaseMirror </strong></font></td>"
    Add-Content $ServicesFileName "</tr>"

        writeTableHeader $ServicesFileName
             
        foreach ($item in $DatabaseMirror)
        {
            writeServiceInfo $ServicesFileName $item.Application $item.Database $item.Pool $item.PrimaryState $item.PrimaryDBServer $item.MirrorDBServer
        }
    Add-Content $ServicesFileName "</table>"
  

writeHtmlFooter $ServicesFileName

Function sendEmail
{ param($from,$to,$subject,$smtphost,$htmlFileName)
$body = Get-Content $htmlFileName
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
$msg.isBodyhtml = $true
$smtp.send($msg)
}

### Configuring Email Parameters
sendEmail LyncDatabaseMirror@domain.com EUCUnifiedComms@domain.com "Lync Database Mirror Report - $Date" SMTPG-GW $ServicesFileName
