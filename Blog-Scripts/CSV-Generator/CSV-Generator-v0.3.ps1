#region Script Initialization
Clear-Host

#Script Variables
$CurrentDate = "{0:yyyy_MM_dd}" -f (get-date)
$DailyCounter = 0
$LogsFolder = "D:\SourceFiles\Scripts\SyntheticTransaction\"
$TargetFQDN = "servername.fqdn"

#Default Tests
$FileName_Test_CsAddressbookservice    = "Test-CsAddressbookservice.csv"    #(ABS)
$FileName_Test_CsAddressBookWebQuery   = "Test-CsAddressBookWebQuery.csv"   #(ABWQ)
$FileName_Test_CsAVConference          = "Test-CsAVConference.csv"          #(AvConference)
$FileName_Test_CsGroupIM               = "Test-CsGroupIM.csv"               #(IM Conferencing)
$FileName_Test_CsIM                    = "Test-CsIM.csv"                    #(P2P IM)
$FileName_Test_CsP2PAV                 = "Test-CsP2PAV.csv"                 #(P2PAV)
$FileName_Test_CsPresence              = "Test-CsPresence.csv"              #(Presence)
$FileName_Test_CsRegistration          = "Test-CsRegistration.csv"          #(Registration)

#Non-Default Tests
$FileName_Test_CsPstnPeerToPeerCall    = "Test-CsPstnPeerToPeerCall.csv"    #(PSTN)
$FileName_Test_CsAVEdgeConnectivity    = "Test-CsAVEdgeConnectivity.csv"    #(AudioVideo EDGE)
$FileName_Test_CsDataConference        = "Test-CsDataConference.csv"        #(DataConference)
$FileName_Test_CsExumConnectivity      = "Test-CsExumConnectivity.csv"      #(ExumConnectivity)
$FileName_Test_CsGroupIMJoinLauncher   = "Test-CsGroupIMJoinLauncher.csv"   #(JoinLauncher)
$FileName_Test_CsMCXP2PIM              = "Test-CsMCXP2PIM.csv"              #(MCXP2PIM)
$FileName_Test_CsPstnOutboundCall      = "Test-CsPstnOutboundCall.csv"      #(PstnOutboundCall)

#Disabled Tests
$FileName_Test_CsASConference          = "Test-CsASConference.csv"          # Currently Disabled
$FileName_Test_CsClientAuth            = "Test-CsClientAuth.csv"            # Currently Disabled
$FileName_Test_CsXmppIM                = "Test-CsXmppIM.csv"                # Currently Disabled
$FileName_Test_CsPersistentChatMessage = "Test-CsPersistentChatMessage.csv" # Currently Disabled
$FileName_Test_CsUnifiedContactStore   = "Test-CsUnifiedContactStore.csv"   # Currently Disabled
#endregion


#region Import Lync Module
Write-Host "Please wait while the Lync PowerShell Module is loading" -ForegroundColor Yellow
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
#endregion

#region Tests Users
$Passwrd = "PASSWORD"
$PasswrdS = $Passwrd | ConvertTo-SecureString -AsPlainText -Force
$CheckUser1    = "netbios_domain\lynctest01"
$CheckUser2    = "netbios_domain\lynctest02"
$GetUser1 = get-csuser -identity $CheckUser1
$GetUser2 = get-csuser -identity $CheckUser2
$Sip1 = $getUser1.sipaddress
$Sip2 = $getUser2.sipaddress
$Creds1 = New-Object System.Management.Automation.PSCredential -ArgumentList $CheckUser1, $PasswrdS
$Creds2 = New-Object System.Management.Automation.PSCredential -ArgumentList $CheckUser2, $PasswrdS
#endregion

Do
{
    #Creating Today's date folder
    New-Item -ItemType Directory -Force -Path $LogsFolder\$CurrentDate

    #Clean items older then 15 days backs (Files & Folders)
    if (!(Test-Path -Path $LogsFolder\$CurrentDate))
    {
        Write-Output "$LogsFolder\$CurrentDate not found, nothing to delete"
    }
    Else
    {
        $Days = "-15" 
        $Today = Get-Date
        $DatetoDelete = $Today.AddDays($Days)
        Get-ChildItem $LogsFolder -recurse | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -force -Recurse | Out-Null
    }

    ###########################################
    #############  DEFAULT TESTS  #############
    ###########################################
    
    # 1) Test-CSAddressbookService
    Write-Host "Starting Test-CSAddressbookService" -ForegroundColor Yellow
    $Results_TestCsAddressbookservice = Test-CSAddressbookService -TargetFQDN $TargetFQDN
    $Results_TestCsAddressbookservice | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsAddressbookservice" -NoTypeInformation -Append

    # 2) Test-CSAddressBookWebQuery
    Write-Host "Starting Test-CSAddressBookWebQuery" -ForegroundColor Yellow
    $Results_TestCsAddressbookWebQuery = Test-CSAddressBookWebQuery -TargetFQDN $TargetFQDN
    $Results_TestCsAddressbookWebQuery | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsAddressBookWebQuery" -NoTypeInformation -Append
    
    # 3) Test-CsCsAVConference    
    Write-Host "Starting Test-CsAVConference" -ForegroundColor Yellow
    $Results_TestCsAVConference = Test-CsAVConference -TargetFqdn $TargetFQDN -SenderSipAddress $Sip1 -ReceiverSipAddress $Sip2
    $Results_TestCsAVConference | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsAVConference" -NoTypeInformation -Append

    # 4) Test-CsGroupIM        
    Write-Host "Starting Test-CsGroupIM" -ForegroundColor Yellow
    $Results_TestCsGroupIM = Test-CsGroupIM -TargetFqdn $TargetFQDN
    $Results_TestCsGroupIM | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsGroupIM" -NoTypeInformation -Append

    # 5) Test-CsIM        
    Write-Host "Starting Test-CsIM" -ForegroundColor Yellow
    $Results_TestCsIM = Test-CsIM -TargetFqdn $TargetFQDN -SenderSipAddress $Sip1 -ReceiverSipAddress $Sip2
    $Results_TestCsIM | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsIM" -NoTypeInformation -Append
    
    # 6) Test-CsP2PAV
    Write-Host "Starting Test-CsP2PAV" -ForegroundColor Yellow
    $Results_TestCsP2PAV = Test-CsP2PAV -TargetFqdn $TargetFQDN -SenderSipAddress $Sip1 -ReceiverSipAddress $Sip2
    $Results_TestCsP2PAV | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsP2PAV" -NoTypeInformation -Append
    
    # 7) Test-CsPresence
    Write-Host "Starting Test-CsPresence" -ForegroundColor Yellow
    $Results_TestCsPresence = Test-CsPresence -TargetFqdn $TargetFQDN
    $Results_TestCsPresence | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsPresence" -NoTypeInformation -Append

    # 8) Test-CsRegistration
    Write-Host "Starting Test-CsRegistration" -ForegroundColor Yellow
    $Results_TestCsRegistration = Test-CsRegistration -TargetFqdn $TargetFQDN -UserSipAddress $Sip1
    $Results_TestCsRegistration | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsRegistration" -NoTypeInformation -Append
    
    ###########################################
    #############  NON DEFAULT TESTS  #############
    ###########################################

    <# 9) Test-CsPstnPeerToPeerCall
    Write-Host "Starting Test-CsPstnPeerToPeerCall" -ForegroundColor Yellow
    $Results_CsPstnPeerToPeerCall = Test-CsPstnPeerToPeerCall -TargetFQDN $TargetFQDN
    $Results_CsPstnPeerToPeerCall | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsPstnPeerToPeerCall" -NoTypeInformation -Append #>

    # 10) Test-CsAVEdgeConnectivity
    Write-Host "Starting Test-CsAVEdgeConnectivity" -ForegroundColor Yellow
    $Results_CsAVEdgeConnectivity = Test-CsAVEdgeConnectivity -TargetFQDN $TargetFQDN
    $Results_CsAVEdgeConnectivity | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsAVEdgeConnectivity" -NoTypeInformation -Append

    # 11) Test-CsDataConference
    Write-Host "Starting Test-CsDataConference" -ForegroundColor Yellow
    $Results_CsDataConference = Test-CsDataConference -TargetFQDN $TargetFQDN
    $Results_CsDataConference | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsDataConference" -NoTypeInformation -Append

    # 12) Test-CsExumConnectivity
    Write-Host "Starting Test-CsExumConnectivity" -ForegroundColor Yellow
    $Results_CsExumConnectivity = Test-CsExumConnectivity -TargetFQDN $TargetFQDN
    $Results_CsExumConnectivity | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsExumConnectivity" -NoTypeInformation -Append

    # 13) Test-CsGroupIMJoinLauncher
    Write-Host "Starting Test-CsGroupIMJoinLauncher" -ForegroundColor Yellow
    $Results_CsGroupIMJoinLauncher = Test-CsGroupIM -TestJoinLauncher -TargetFQDN $TargetFQDN
    $Results_CsGroupIMJoinLauncher | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsGroupIMJoinLauncher" -NoTypeInformation -Append

    # 14) Test-CsMCXP2PIM 
    Write-Host "Starting Test-CsMCXP2PIM" -ForegroundColor Yellow
    $Results_CsMCXP2PIM = Test-CsMCXP2PIM -TargetFQDN $TargetFQDN -SenderSipAddress $Sip1 -ReceiverSipAddress $Sip2 -Authentication Negotiate -SenderCredential $Creds1 -ReceiverCredential $Creds2
    $Results_CsMCXP2PIM | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsMCXP2PIM " -NoTypeInformation -Append

    # 15) Test-CsPstnOutboundCall 
    Write-Host "Starting Test-CsPstnOutboundCall" -ForegroundColor Yellow
    $Results_CsPstnOutboundCall = Test-CsPstnOutboundCall -TargetFQDN $TargetFQDN -UserSipAddress $Sip1 -TargetPstnPhoneNumber "+1515131720"
    $Results_CsPstnOutboundCall | Select-Object *,@{Name="TimeStamp";Expression={Get-Date}} | Export-Csv -Path `
    "$LogsFolder\$CurrentDate\$FileName_Test_CsPstnOutboundCall " -NoTypeInformation -Append


    $DailyCounter += 1
    Start-Sleep 120
}

While ($DailyCounter -lt 350)
