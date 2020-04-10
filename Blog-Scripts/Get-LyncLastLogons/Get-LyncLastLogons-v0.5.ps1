<#  
.SYNOPSIS  
	Shows users last registered to a Lync pool and the type of client it registered with

.NOTES  
  Version      				: 0.5
  Rights Required			: Local admin on server and SQL (Run PowerShell as Administrator)
  Lync Version				: 2013 (tested on March 2014 CU Update)
  Authors       			: Guy Bachar, Yoav Barzilay
  Last Update               : 5-July-2014
  Twitter/Blog	    : @GuyBachar, http://guybachar.us
  Twitter/Blog	    : @y0avb, http://y0av.me

  Source   : http://blogs.technet.com/b/meacoex/archive/2011/07/19/list-connections-and-users-connected-to-lync-registrar-pool.aspx
  Source   : http://blogs.technet.com/b/nexthop/archive/2011/03/10/list_2d00_users_2d00_and_2d00_endpoints_2d00_direct.aspx
  Source   : http://www.ehloworld.com/269
  Source   : http://blogs.technet.com/b/dodeitte/archive/2011/05/11/how-to-get-the-last-time-a-user-registered-with-a-front-end.aspx

.VERSION
  0.41 - Added Export Options & Display Options
  0.5 - Added Support for Get-CsUserPoolInfo Stats to identify the User Primary Server within a Pool and its associated Backup Servers
	
#>

#region Script Initialization
Clear-Host
$Date = Get-Date -DisplayHint DateTime
#endregion

#region Script Information
Write-Host "-------------------------------------------------------" -BackgroundColor DarkGreen
Write-Host
Write-Host "Lync last logons" -ForegroundColor Green
Write-Host "Version: 0.5" -ForegroundColor Green
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
Write-Host "Please wait while we're loading Lync PowerShell Module..."
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
Write-Host    
Write-Host "Done!" -ForegroundColor Green
Start-Sleep -Seconds 1
#endregion

#region Setting Parameters for Display or Export
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
#endregion

#region Searching for Lync Pool Registrars
$pool=Get-CsPool | Where-Object {$_.Services -like "*Registrar*"}
Write-Host 
Write-Host "****************************************************"
Write-Host
Write-Host "Searching for Registrars..." -ForegroundColor Yellow
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
#endregion

<#region Searching for Lync Servers within Pool
$servers = $pool.Computers
Write-Host 
Write-Host "****************************************************"
Write-Host
Write-Host "Searching for Computers within the Pool..." -ForegroundColor Yellow
Write-Host "Done!" -foregroundcolor Green
Write-Host
if ($servers.Count -eq 1)
    {
        Write-Host "There is Only one server in this pool" -ForegroundColor DarkCyan
        Write-Host
		Write-Host "Server name is $servers.fqdn"
        Write-Host
        $servers = $servers[0]

    }
elseIf ($servers.Count -gt 1)
	{
		Write-Host "Servers Selection" -ForegroundColor DarkCyan
        Write-Host
		Write-Host "Server No.    Server Name"
        Write-Host
	for ($i=0; $i -lt $servers.Count; $i++)
		{
			$a = $i + 1
			Write-Host ($a, $servers[$i]) -Separator "                "
		}
    $Range = '(1-' + $servers.Count + ')'
		Write-Host
		$Select = Read-Host "Please choose the Server you want to query with active users" $range
		$Select = $Select - 1
		If (($Select -gt $servers.Count-1) -or ($Select -lt 0))
		{
			Write-Host "The only options available are"$range, Please try again. -ForegroundColor Red
			Exit
		}
        Else
		{
			$servers = $servers[$Select]
		}
#>#endregion

#region Create empty variable that will contain the user registration records
Write-Host Searching for active users on this pool... -ForegroundColor Yellow
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
   		
        
        try
        {
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
        }
        catch
        { write-host "Error Conencting to local SQL service on $ComputerName, Please verify connectivity and permissions" -ForegroundColor Red }
     
        #Append the results to the reuslts from the previous servers
        $overallrecords = $overallrecords + $Results.Tables[0] 
}
#endregion

$filedate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$ServicesFileName = $env:TEMP+"\LastLogonExport-"+$filedate+".csv"
$ListUsers = @()

#Filter Dates
#$DateToCompare = (Get-date).AddDays(-1)
#$overallrecords = $overallrecords | Where-Object {$_.LastRegisterTime -gt $DateToCompare}

$overallrecords | ForEach-Object{ 

    # save a reference to the current user
    $user = $_           
     
    # get the user registrar info
    $UserInfo = Get-CsUser -Identity $user.UserName -Filter {RegistrarPool -ne $null} -ErrorAction SilentlyContinue | 
    Get-CsUserPoolInfo | Where {$_.PrimaryPoolFQDN -eq $Pool.Fqdn} | 
    Select-Object Identity,PrimaryPoolFQDN,BackupPoolFQDN,PrimaryPoolPrimaryRegistrar,PrimaryPoolBackupRegistrars

    # comment out to add just the LastRegisterTime property
    $user | Add-Member -MemberType NoteProperty -Name "Primary Pool FQDN" -Value $UserInfo.PrimaryPoolFQDN -ErrorAction SilentlyContinue
    $user | Add-Member -MemberType NoteProperty -Name "Main Server" -Value $UserInfo.PrimaryPoolPrimaryRegistrar -ErrorAction SilentlyContinue
    $user | Add-Member -MemberType NoteProperty -Name "Backup Servers" -Value $UserInfo.PrimaryPoolBackupRegistrars -ErrorAction SilentlyContinue
    
    $ListUsers = $ListUsers + $user
    
} 

If ($UserInput -eq 1)    { 
                            $ListUsers |Select UserName,ClientVersion,LastRegisterTime,ExpiresAt,"Primary Pool FQDN","Main Server","Backup Servers"  | Export-Csv -Path $ServicesFileName
                            Write-Host
                            Write-Host "The File is located under $ServicesFileName" }
elseif ($UserInput -eq 2)     {$ListUsers | Select UserName,ClientVersion,LastRegisterTime,ExpiresAt,"Primary Pool FQDN","Main Server","Backup Servers" | Out-GridView}
else { $ListUsers | Select UserName,ClientVersion,LastRegisterTime,ExpiresAt,"Primary Pool FQDN","Main Server","Backup Servers" | Format-Table -AutoSize}
