<#  
.SYNOPSIS  
    

.NOTES  
    Version                   : 0.5
    Rights Required           : Local admin
    Lync Version              : 2013 (tested on March 2015 Update)
    Skype4B Version           : 2015 (tested on RTM)
    Authors                   : Guy Bachar
    Last Update               : 8-June-2015
    Twitter/Blog              : @GuyBachar, http://guybachar.wordpress.com
    Twitter/Blog              : @y0avb, http://y0av.me


.VERSION
    0.1 - Initial Version for reporting for Lync/Skype4B
    0.2 - Changes in Query for Weekly stats instead of daily
    0.3 - Initial Version for reporting Lync/Skype4B Users which reported with Poor Call quality Calls
    0.4 - Updating Poor Call state per CQM
    0.5 - Adding parameters for running remotely against a Front End Server and validation checks
      
#>

param(
[Parameter(Position=0, Mandatory=$false)]
$To,
[Parameter(Position=1, Mandatory=$false)]
$From,
[Parameter(Position=2, Mandatory=$false)]
$SmtpServer,
[Parameter(Position=3, Mandatory=$False) ][ValidateNotNullorEmpty()][switch] $RemoteConnection
)


function WriteError ($Message) {
	$a = (Get-Host).UI.RawUI
	$ORIColor = $a.ForegroundColor
	$a.ForegroundColor = "red"
	WRITE-HOST "----------------------------------------------------------------------------"
	WRITE-HOST $Message
	WRITE-HOST "----------------------------------------------------------------------------"
	$a.ForegroundColor = $OriColor
}

function WriteTitle ($TitleMessage) {
	$a = (Get-Host).UI.RawUI
	$ORIColor = $a.ForegroundColor
	$a.ForegroundColor = "green"
	WRITE-HOST "----------------------------------------------------------------------------"
	WRITE-HOST $TitleMessage
	WRITE-HOST "----------------------------------------------------------------------------"
	$a.ForegroundColor = $OriColor
}

function GetMonitoringDatabases(){
	$SqlTable = new-object System.Data.DataTable "SqlTable"
	$col1 = New-Object system.Data.DataColumn SQLInstance,([String])
	$col2 = New-Object system.Data.DataColumn LyncVersion,([String])
	$col3 = New-Object system.Data.DataColumn YesNo,([String])

	#Add the Columns
	$SqlTable.columns.add($col1)
	$SqlTable.columns.add($col2)
	$SqlTable.columns.add($col3)

	$CurVerbose = $VerbosePreference
	$VerbosePreference = "SilentlyContinue"

	#Get Monitoring Databases
	$MonStores = Get-CsService -MonitoringDatabase
	foreach ($Store in $MonStores) {

		$RegPoolFqdn = $NULL
		if ($Store.Version -eq 5) {
			# This Lync 2010 database, go via MonitoringStore
			$DepMonServer = Get-CsService -id $Store.DependentServiceList[0]
			if ($DepMonServer -ne $NULL) {
				# Check if the Monitoring Server has any Registrars depending on it
				$DepRegistrar = $DepMonServer.DependentServiceList[0]
				if ($DepRegistrar -ne $NULL) {
					$RegPoolFqdn = $DepRegistrar.Replace("Registrar:","")
				}
			}
		}
		else {
			# This is 2013 or SfB2015 monitoring database
			# Get first dependant registrar - Registrar:pool1.contoso.com
			$DepRegistrar = $Store.DependentServiceList[0]
			if ($DepRegistrar -ne $NULL) {
				$RegPoolFqdn = $DepRegistrar.Replace("Registrar:","")
			}
		}

		if (($OurPSVersion -gt 4) -and ($RegPoolFqdn -ne $NULL)) {
			# This Lync 2013 or later management shell with Get-CsDatabaseMirrorState command
			# see if there is a mirror relationship for this monitoring database
			$MirrorStates = Get-CsDatabaseMirrorState -PoolFqdn $RegPoolFqdn -DatabaseType Monitoring -ErrorAction SilentlyContinue
		}
		else {
			$MirrorStates = $NULL
		}
		

		if ($MirrorStates -eq $NULL) {
			# if there is no mirror relationship save SQL instance information
			#Create a row
			$row = $SqlTable.NewRow()
			# only use it if one or more registrars depends on this monitoring database
			if ($RegPoolFqdn -ne $NULL) {
				$row.YesNo = "Yes"
			}
			else {
				$row.YesNo = "No"
			}
			if ($Store.Version -eq 7) {
				$row.LyncVersion = "SfB2015"
			}
			if ($Store.Version -eq 6) {
				$row.LyncVersion = "Lync2013"
			}
			if ($Store.Version -eq 5) {
				$row.LyncVersion = "Lync2010"
			}
			$row.SqlInstance = $Store.PoolFqdn + "\" + $Store.SqlInstanceName
			#Add the row to the table
			$SqlTable.Rows.Add($row)
		}
		else {
			# look for state for LcsCDR
			foreach ($db in $MirrorStates) {
				if ($db.DatabaseName.ToLower() -eq "lcscdr") {

					# Get monitoring database FQDN from Registrar Info
					$RegInfo = Get-CsService -Registrar -PoolFqdn $RegPoolFqdn

					# Only create row if state is Principal and we are having that SQL server in the loop
					if ($db.StateOnPrimary -eq "Principal") {

						# CDR is on primary
						# See if we are looing at the primary store right now
						$tmpStr = $RegInfo.MonitoringDatabase.Replace("MonitoringDatabase:","")
						if ($tmpStr -eq $Store.PoolFqdn) {

							#Create a row
							$row = $SqlTable.NewRow()
							$row.YesNo = "Yes"
							$MonInfo = Get-CsService -MonitoringDatabase -PoolFqdn $tmpStr
							$row.SqlInstance = $tmpStr + "\" + $MonInfo.SqlInstanceName

							if ($Store.Version -eq 7) {
								$row.LyncVersion = "SfB2015"
							}
							if ($Store.Version -eq 6) {
								$row.LyncVersion = "Lync2013"
							}
							if ($Store.Version -eq 5) {
								$row.LyncVersion = "Lync2010"
							}

							#Add the row to the table
							$SqlTable.Rows.Add($row)
						}
                    			}
					else {
						# CDR is on Mirror
						# See if we are looing at the mirror store right now, if so create row
						$tmpStr = $RegInfo.MirrorMonitoringDatabase.Replace("MonitoringDatabase:","")
						if ($tmpStr -eq $Store.PoolFqdn) {

							#Create a row
							$row = $SqlTable.NewRow()
							$row.YesNo = "Yes"
							$MonInfo = Get-CsService -MonitoringDatabase -PoolFqdn $tmpStr
							$row.SqlInstance = $tmpStr + "\" + $MonInfo.SqlInstanceName

							if ($Store.Version -eq 7) {
								$row.LyncVersion = "SfB2015"
							}
							if ($Store.Version -eq 6) {
								$row.LyncVersion = "Lync2013"
							}
							if ($Store.Version -eq 5) {
								$row.LyncVersion = "Lync2010"
							}

							#Add the row to the table
							$SqlTable.Rows.Add($row)
						}

					}
					
				}
			}
		}
	}

	$VerbosePreference = $CurVerbose
	return $SqlTable
}

function New-PoorCallReport {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $PoorCallEntry	
        )
	begin {
$css = @'
	<style type="text/css">
	body { font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif;}
	table {border-collapse: separate; background-color: #e6edef; border: 3px solid #103E69; caption-side: bottom;}
	td { border:1px solid #103E69; margin: 3px; padding: 3px; vertical-align: top; background: #e6edef; color: #000;font-size: 12px;}
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
		[void]$sb.AppendLine("<tr><td colspan='2'><strong><center>Users Reported with Poor Call for $StartDate - $EndDate</cenetr></strong></td></tr>")
		[void]$sb.AppendLine("<tr>")
		#[void]$sb.AppendLine("<td><strong>Call Date</strong></td>")
		[void]$sb.AppendLine("<td><strong>Caller Uri (Source User/Number)</strong></td>")
		[void]$sb.AppendLine("<td><strong>Reported Number of Poor Calls</strong></td>")
		[void]$sb.AppendLine("</tr>")
	}
	
	process {
		[void]$sb.AppendLine("<tr>")
		#[void]$sb.AppendLine("<td>$($PoorCallEntry.CallDate)</td>")
		[void]$sb.AppendLine("<td>$($PoorCallEntry.CallerURI)</td>")
		[void]$sb.AppendLine("<td>$($PoorCallEntry.NumberOfPoorCalls)</td>")
		[void]$sb.AppendLine("</tr>")
		$cmdletparameters = $null
	}
	
	end {
		[void]$sb.AppendLine("</table>")
		Write-Output $sb.ToString()
	}
}


#
# Main Start
#

#region Check PowerShell Version
$tmpStrPowerShell = "Confirming PowerShell Version..."
WriteTitle $tmpStrPowerShell
if ($PSVersionTable.PSVersion.Major -lt 3.0)
{
Write-Host 
Write-Host PowerShell Version is not compatible with this script. -ForegroundColor Red
Write-Host You are running version $PSVersionTable.PSVersion.Major -ForegroundColor Red
Write-Host Version must be at least 3.0. -ForegroundColor Red
Write-Host
}
#endregion


#region Connect to Server
if ($RemoteConnection.IsPresent)
{
    Write-Host "Waiting for Credentials..." -ForegroundColor Yellow
    $cred = Get-Credential -Message "Please enter your Lync Administrator's credentials:" 
    Write-Host
    $LyncServer = Read-Host "Please Enter the FQDN of your Lync Server or pool"
    Write-Host Connecting to Pool $LyncServer... -ForegroundColor Yellow
    $sessionOption = New-PSSessionOption -SkipRevocationCheck
    $session = New-PSSession -ConnectionURI https://$LyncServer/OcsPowershell -Credential $cred -SessionOption $sessionOption -ErrorAction Stop
    Import-PsSession $session -AllowClobber
    Clear-Host
    Write-Host
    Write-Host Connected to $session.ComputerName -ForegroundColor Yellow
}
#endregion

else 
{
    #Test if Lync or SkypeForBusiness modules are loaded
    $LyncLoaded = Get-Module -Name Lync
    $SfBLoaded = Get-Module -Name SkypeForBusiness
    if (($LyncLoaded -eq $NULL) -and ($SfBLoaded -eq $NUll)) {
        WriteError "Please run this script in Lync or Skype for Business Server Management Shell"
        Exit
    }
    if ($LyncLoaded -ne $NULL) {
	    $OurPSVersion = (Get-Module -Name Lync).Version.Major
    }
    if ($SfBLoaded -ne $NULL) {
	    $OurPSVersion = (Get-Module -Name SkypeForBusiness).Version.Major
    }
}


# Get Monitoring Databases and Pools to collect from
$MonInstances = GetMonitoringDatabases
$MonInstances | Foreach-Object {$tmpStr = $_.SQLInstance + ", " + $_.LyncVersion + ", " + $_.YesNo; Write-Verbose $tmpStr}

# get SIP domains
$CurVerbose = $VerbosePreference
$VerbosePreference = "SilentlyContinue"
$SipDomains = Get-CsSipDomain
$VerbosePreference = $CurVerbose
$DefaultSipDomain = ($SipDomains | Where {$_.IsDefault -eq "True"}).Identity
$strResults = $NULL
$date = (get-date -Format MM.dd.yyyy).toString()
$StartDate = (get-date -Format d (get-date).AddDays(-7)).toString()
$EndDate   = (get-date -Format d).toString()

$SQLQuery = "USE QoEMetrics;  `
DECLARE @beginTime AS DateTime = '$StartDate'; `
DECLARE @endTime   AS DateTime = '$EndDate'; `
SELECT `
	--CONVERT(date,s.StartTime) AS CallDate `
	CallerUser.URI AS CallerUri `
	--,COUNT(CallerUser.URI) AS NumberOfPoorCalls `
	,COUNT(DISTINCT s.StartTime) AS NumberOfPoorCalls `
FROM [Session] s WITH (NOLOCK) `
	INNER JOIN [MediaLine] AS m WITH (NOLOCK) ON  `
		m.ConferenceDateTime = s.ConferenceDateTime `
		AND m.SessionSeq = s.SessionSeq			 `
	INNER JOIN [AudioStream] AS a WITH (NOLOCK) ON `
		a.MediaLineLabel = m.MediaLineLabel     `
		and a.ConferenceDateTime = m.ConferenceDateTime  `
		and a.SessionSeq = m.SessionSeq `
	INNER JOIN [User] AS CallerUser WITH (NOLOCK) ON `
		CallerUser.UserKey = s.CallerURI `
	INNER JOIN [User] AS CalleeUser WITH (NOLOCK) ON `
		CalleeUser.UserKey = s.CalleeURI `
	LEFT JOIN [NetworkConnectionDetail] AS CallerNcd WITH (NOLOCK) ON  `
		CallerNcd.NetworkConnectionDetailKey = m.CallerNetworkConnectionType  `
	LEFT JOIN [NetworkConnectionDetail] AS CalleeNcd WITH (NOLOCK) ON  `
		CalleeNcd.NetworkConnectionDetailKey = m.CalleeNetworkConnectionType `
	LEFT JOIN [Endpoint] AS CallerEndpoint  `
              	ON s.CallerEndpoint = CallerEndpoint.EndpointKey  `
	LEFT JOIN [Endpoint] AS CalleeEndpoint  `
              ON s.CalleeEndpoint = CalleeEndpoint.EndpointKey  `
	LEFT JOIN [Device] AS CallerCaptureDevice  `
              ON m.CallerCaptureDev = CallerCaptureDevice.DeviceKey  `
	LEFT JOIN [Device] AS CallerRenderDevice  `
              ON m.CallerRenderDev = CallerRenderDevice.DeviceKey  `
	LEFT JOIN [Device] AS CalleeCaptureDevice  `
              ON m.CalleeCaptureDev = CalleeCaptureDevice.DeviceKey  `
	LEFT JOIN [Device] AS CalleeRenderDevice  `
              ON m.CalleeRenderDev = CalleeRenderDevice.DeviceKey  `
WHERE `
	s.StartTime >= (@beginTime) and s.StartTime < (@endTime) `
	and CallerUser.URI like '%@$DefaultSipDomain%' `
	and ((PacketLossRate > 0.1 OR DegradationAvg > 1.0 OR RoundTrip > 500 OR JitterInterArrival > 30 OR RatioConcealedSamplesAvg > 0.07)) `
GROUP BY CallerUser.URI, CONVERT (date,s.StartTime) `
HAVING (COUNT(DISTINCT s.StartTime) > 1) `
ORDER  BY NumberOfPoorCalls DESC" 

# Run through all Monitoring database and collect data sourced from LcsCDR
foreach ($MonDb in $MonInstances) {

    if ($MonDb.YesNo -eq "Yes") {
	    $CDRQoEInstance = $MonDb.SqlInstance
	    $LyncVersion = $MonDb.LyncVersion
        
	    $tmpStr = "Collecting Reporting Metrics from " + $CDRQoEInstance + " databases"
	    WriteTitle $tmpStr
        
        # Log information
	    $alltext= "RunDate: " + $date + "`r`nQoEInstance: "+$CDRQoEInstance+"`r`nLyncVersion: "+$LyncVersion+"`r`nMonthlyActiveVoiceUsers: "+($MAVU.ActiveUser | measure-object).Count+"`r`nMonthlyActiveUsers: "+($MAU.ActiveUser | measure-object).Count
    
    
        #Connection String
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server = $CDRQoEInstance; Database = QoEMetrics; Integrated Security = True"
    
        #Define SQL Command     
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SQLQuery
        $SqlCmd.Connection = $SqlConnection

        #Get the results
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd        
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
        $SqlConnection.Close()

        #Append All strResults
        $strResults = $strResults + $DataSet.Tables[0]
     
    }
}


if (($From.Length -gt 0) -AND ($To.Length -gt 0) -AND ($SmtpServer.Length -gt 0))
{
    Send-MailMessage -To $To `
    -From $From `
    -Subject "Reported Users with Poor Calls for $StartDate - $EndDate" `
    -Body ($strResults | New-PoorCallReport) `
    -SmtpServer $SmtpServer `
    -BodyAsHtml
}

else
{
    $FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
    $Report = $strResults | New-PoorCallReport
    $Report | Out-File $env:TEMP"\UsersPoolCallReport-"$FileDate".html"
    Invoke-Item $env:TEMP"\UsersPoolCallReport-"$FileDate".html"
}