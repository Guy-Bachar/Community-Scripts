<#  
.SYNOPSIS  
	Shows users last registered to a Lync pool and the type of client it registered with

.NOTES  
  Version      				: 0.4.1
  Rights Required			: Local admin on server and SQL (Run PowerShell as Administrator)
  Lync Version				: 2013 (tested on March 2014 CU Update)
  Author       				: Guy Bachar
  Twitter/Blog	    : @GuyBachar, http://guybachar.us
  Twitter/Blog	    : @y0avb, http://y0av.me

  Source    : http://blogs.technet.com/b/meacoex/archive/2011/07/19/list-connections-and-users-connected-to-lync-registrar-pool.aspx
  Source    : http://blogs.technet.com/b/nexthop/archive/2011/03/10/list_2d00_users_2d00_and_2d00_endpoints_2d00_direct.aspx
  Source   : http://www.ehloworld.com/269
  Source   : http://blogs.technet.com/b/dodeitte/archive/2011/05/11/how-to-get-the-last-time-a-user-registered-with-a-front-end.aspx

.Version
  0.41 - Added Export Options & Display Options

	
#>

Clear-Host
Write-Host "-------------------------------------------------------"
Write-Host
Write-Host "Lync last logons"
Write-Host
Write-Host "Version: 0.4.1"
Write-Host 
Write-Host "Authors:"
Write-Host
Write-Host " Guy Bachar        @GuyBachar     http://guybachar.us"
Write-Host " Yoav Barzilay     @y0avb         http://y0av.me"
Write-host
$Date = Get-Date -DisplayHint DateTime
Write-Host "-------------------------------------------------------"
Write-Host
Write-Host "Remember: Please run this script as Administrator" -ForegroundColor Black -BackgroundColor Yellow
Write-Host


# Setting Parameters for Display or Export
Write-Host "****************************************************"
Write-Host "Output Selection (Default is output to screen)" -ForegroundColor DarkCyan
Write-Host
Write-Host "1)     Export to CSV"
Write-Host "2)     Export to GridView"
Write-Host "3)     Export to Screen"
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


# Searching for Lync Registrars
$pool=Get-CsPool | Where-Object {$_.Services -like "*Registrar*"}
Write-Host 
Write-Host "****************************************************"
Write-Host
Write-Host "Searching for Registraras..." -ForegroundColor Yellow
Write-Host "Done!" -foregroundcolor Green
Write-Host
If ($pool.Count -gt 1)
	{
		Write-Host "Registrar Selection" -ForegroundColor DarkCyan
        Write-Host
		Write-Host "Registrar No.    Registrar Name"
        Write-Host
	for ($i=0; $i -lt $Pool.Count; $i++)
		{
			$a = $i + 1
			Write-Host ($a, $Pool[$i].Identity) -Separator "                "
		}

}
$Range = '(1-' + $Pool.Count + ')'
		Write-Host
		$Select = Read-Host "Please choose the Registrar you want to query" $range
		$Select = $Select - 1
		If (($Select -gt $Pool.Count-1) -or ($Select -lt 0))
		{
			Write-Host "The only options available are"$range, Please try again. -ForegroundColor Red
			Exit
		}
		Else
		{
			$Pool = $Pool[$Select]
		}

#Create empty variable that will contain the user registration records
$overallrecords = $null

#Loop through a front end computers in the pool
Foreach ($Computer in $Pool.Computers){

    #Get computer name from fqdn
    $ComputerName = $Computer.Split(".")[0]

    #Defined Connection String
    $connstring = "server=$ComputerName\rtclocal;database=rtcdyn;`
        trusted_connection=true;"

    #Define SQL Command     
    $command = New-Object System.Data.SqlClient.SqlCommand
    $command.CommandText = "Select
        R.UserAtHost as UserName, `
        HRD.LastNewRegisterTime as LastRegisterTime, `(cast (RE.ClientApp as `
        varchar (100))) as ClientVersion, `
        EP.ExpiresAt, `
        '$computer' as RegistrarFQDN `
        From `
       rtcdyn.dbo.RegistrarEndpoint RE `
       Inner Join `
        rtc.dbo.Resource R on R.ResourceId = RE.OwnerId `
       Inner Join `
        rtcdyn.dbo.Endpoint EP on EP.EndpointId = RE.EndpointId `
	   Inner Join `
	    rtcdyn.dbo.HomedResourceDynamic HRD on HRD.OwnerId = R.ResourceId `
        Order By UserName, ClientVersion "
		
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
}

$filedate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LastLogonExport-"+$filedate+".csv"

If ($UserInput -eq 1)    { 
                            $overallrecords | Export-Csv -Path $ServicesFileName
                            Write-Host
                            Write-Host "The File is located under $ServicesFileName" }
elseif ($UserInput -eq 2)     {$overallrecords | Out-GridView}
else { $overallrecords }