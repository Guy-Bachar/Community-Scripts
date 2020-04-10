<#  
.SYNOPSIS  
	This script shows Lync users last logon to a Lync pool based on the Lync CDR database and will display Lync Orphaned Users

.NOTES  
  Version      				: 0.2
  Rights Required			: Local admin on server and SQL (Run PowerShell as Administrator)
  Lync Version				: 2013 (tested on March 2014 CU Update)
  Authors       			: Guy Bachar, Yoav Barzilay
  Last Update               : 8-July-2014
  Twitter/Blog	            : @GuyBachar, http://guybachar.us
  Twitter/Blog	            : @y0avb, http://y0av.me


.VERSION
  0.1 - Initial Version for connecting LCSCDR database
  0.2 - Getting Monitoring Database Automatically
	
#>

#region Script Information
Clear-Host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Lync last logons - LCSCDR Version" -ForegroundColor Green
Write-Host "Version: 0.2" -ForegroundColor Green
Write-Host 
Write-Host "Authors:" -ForegroundColor Green
Write-Host " Guy Bachar    | @GuyBachar | http://guybachar.us" -ForegroundColor Green
Write-Host " Yoav Barzilay | @y0avb     | http://y0av.me" -ForegroundColor Green
Write-host
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
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

#region Setting Parameters for Display or Export
Write-Host
Write-Host "****************************************************"
Write-Host
Write-Host "Output Selection (Default is output to screen)" -ForegroundColor Cyan
Write-Host
Write-Host "1     Export to CSV"
Write-Host "2     Export to GridView"
Write-Host "3     Export to Screen"
Write-Host
$UserInput = Read-Host "Please Enter your choice"
Write-Host
switch ($UserInput) 
    { 
        1 {"You chose: Export to CSV"} 
        2 {"You chose: Export to GridView"} 
        3 {"You chose: Export to Screen"}
        default {"No option selected. Results will be displayed on the screen."}
    }
Write-Host
#endregion

#region Lync Monitoring Server Input
Write-Host "****************************************************"
Write-Host
Write-Host "Monitoring Server and Instances Selection:" -ForegroundColor Cyan
Write-Host
   
Do
    {
        $Error.Clear()
        try
        {
            $MonServer = Get-CsService | Where-Object {$_.Role -eq "MonitoringDatabase"}
    
            If ($MonServer.Count -gt 0)
	        {
		    Write-Host "Server No.   Monitoring Server Name"
            Write-Host
	        for ($i=0; $i -lt $MonServer.Count; $i++)
		        {
			        $a = $i + 1
			        Write-Host ($a, $MonServer[$i].PoolFqdn) -Separator "     "
		        }
            }
            $Range = '(1-' + $MonServer.Count + ')'
		    Write-Host
		    $Select = Read-Host "Please choose Monitoring Server" $range
		    $Select = $Select - 1
		
            If (($Select -gt $MonServer.Count-1) -or ($Select -lt 0))
		    {
			    Write-Host "The only options available are"$range, Please try again. -ForegroundColor Red
			    Exit
		    }
   		    Else
		    {
			    $MonServer = $MonServer[$Select]
            }
            
            Write-Host 
            Write-Host
            Write-Host "Searching for Instances..." -ForegroundColor Yellow
            $SQLReportingServerInstances = [System.Data.Sql.SqlDataSourceEnumerator]::Instance.GetDataSources() | ? { $_.servername -eq $Monserver.PoolFqdn.Split(".")[0]} -ErrorAction Stop
            Write-Host "Done!" -foregroundcolor Green
            Write-Host
            Write-Host "****************************************************"
            #Write-Host
            if ($SQLReportingServerInstances.InstanceName -eq [System.DBNull]::Value)
            {
                $InstanceParameter = "NonRequired"
            }
            elseIf ($SQLReportingServerInstances.Count -gt 1)
	        {
		        #Write-Host "Instances Selection" -ForegroundColor DarkCyan
                Write-Host
		        Write-Host "Instance No.    Instance Name"
                Write-Host
	            for ($i=0; $i -lt $SQLReportingServerInstances.Count; $i++)
		        {
			        $a = $i + 1
			        Write-Host ($a, $SQLReportingServerInstances[$i].InstanceName) -Separator "                "
	    	    }
                $Range = '(1-' + $SQLReportingServerInstances.Count + ')'
                Write-Host
                $Select = Read-Host "Please choose the instance you want to query" $range
                $Select = ($Select - 1)
                If ((($Select -gt $SQLReportingServerInstances.Count-1) -or ($Select -lt 0)) -AND ($InstanceParameter -ne "NonRequired"))
                {
                Write-Host "The only options available are"$range, Please try again. -ForegroundColor Red
                Exit
                }
                Else
                {
                $SQLReportingServerInstances = $SQLReportingServerInstances[$Select].InstanceName
                }
             }
            #$SQLReportingServerInstances | Select-Object InstanceName -ErrorAction Stop
	}
	Catch [Exception]
	{
	    Write-Host "Server " -nonewline; Write-Host $MonServer -foregroundcolor red -nonewline; " does not exist or cannot be contacted. Please try again."
        Write-Host
	}
} 

until (($MonServer -ne $Null) -and (!$Error))

<#List Databases Out of SQL Instance
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
$s = New-Object ('Microsoft.SqlServer.Management.Smo.Server') "$SQLReportingServer\$SQLReportingServerInstances"
$dbs=$s.Databases
$dbs | Get-Member -MemberType Property #>
#endregion

#region Create empty variable that will contain the user registration records
Write-Host
Write-Host Searching for active users on this instance... -ForegroundColor Yellow
$overallrecords = $null

$SqlQuery = "SELECT Users.UserId,UserUri,LastLogInTime,LastConfOrganizedTime,LastCallFailureTime,LastConfOrganizerCallFailureTime `
                        FROM [LcsCDR].[dbo].[UserStatistics] `
                        INNER JOIN [LcsCDR].[dbo].[Users] ON [LcsCDR].[dbo].[Users].UserId = [LcsCDR].[dbo].[UserStatistics].UserId `
                        WHERE LastLogInTime IS NOT NULL `
                        ORDER BY [LcsCDR].[dbo].[UserStatistics].LastLogInTime desc"

#Defnie SQL Connection
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SQLServer = $Monserver.PoolFqdn.Split(".")[0]

if ($InstanceParameter -eq "NonRequired")
        {
            $SqlConnection.ConnectionString = "Server = $SQLServer; Database = lcscdr; Integrated Security = True"
		}else
        {
            $SqlConnection.ConnectionString = "Server = $SQLServer\$SQLReportingServerInstances; Database = lcscdr; Integrated Security = True"
		}
        #Write-Host $SqlConnection.ConnectionString 

try   {
        #Define SQL Command     
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection

        #Get the results
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
        
        $SqlConnection.Close()
        Write-Host "Done!" -foregroundcolor Green
      }
catch {
        write-host "Error Conencting to local SQL service on $SQLReportingServer with instance $SQLReportingServerInstances, Please verify connectivity and permissions" -ForegroundColor Red
        Break
      }
 
#Append the results to the reuslts from the previous servers
$overallrecords = $DataSet.Tables[0]
#endregion

#region How Many Days since last logged in?
Write-Host "****************************************************"
Write-Host
#Write-Host "Days Selection" -ForegroundColor DarkCyan
Write-Host
Do
{
    $Error.Clear()
    $UserDaysInput = Read-Host "Please Provide with amount of days since last logon to display"

    $DateToCompare = (Get-date).AddDays(-$UserDaysInput)
    $overallrecords = $overallrecords | Where-Object {$_.LastLogInTime -lt $DateToCompare} -ErrorAction Ignore
}
until (($UserDaysInput -gt 0) -and (!$Error))
#endregion

#region Script Output Display
$filedate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LastLogonExport-"+$filedate+".csv"
$ListUsers = @()
$overallrecords | ForEach-Object{ 

    # save a reference to the current user
    $user = $_           

    $tspan=New-TimeSpan $user.LastLogInTime (Get-Date);
    $diffDays=($tspan).days;

    # comment out to add just the LastRegisterTime property
    $user | Add-Member -MemberType NoteProperty -Name "Days Since Last Login" -Value ($diffDays)
    $ListUsers = $ListUsers + $user
} 

If ($UserInput -eq 1)
    { 
        $overallrecords| Select-Object UserUri,LastLogInTime,"Days Since Last Login",LastConfOrganizedTime,LastCallFailureTime,LastConfOrganizerCallFailureTime | Export-Csv -Path $ServicesFileName
        Write-Host
        Write-Host "The File is located under $ServicesFileName"
    }
elseif ($UserInput -eq 2)
    {
        $overallrecords | Select-Object UserUri,LastLogInTime,"Days Since Last Login",LastConfOrganizedTime,LastCallFailureTime,LastConfOrganizerCallFailureTime | Out-GridView
    }
else{
        $overallrecords | Select-Object UserUri,LastLogInTime,"Days Since Last Login",LastConfOrganizedTime,LastCallFailureTime,LastConfOrganizerCallFailureTime | Format-Table -AutoSize
    }
#endregion