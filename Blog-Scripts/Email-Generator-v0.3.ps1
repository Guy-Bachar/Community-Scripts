#region Script Initialization
Clear-Host
$CurrentDate = "{0:yyyy_MM_dd}" -f (get-date)
$FileDate = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$EmailFileName = $env:TEMP+"\LyncSyntheticTransaction-"+$FileDate+".htm"
$CSVFolder = "\\SERVERNAME\D$\SourceFiles\Scripts\SyntheticTransaction\$CurrentDate"
$LogsFolder = "D:\SourceFiles\Scripts\SyntheticTransaction"
$PerFailureThresholdWarning = 0.05
$PerFailureThresholdError = 0.1

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

#### Building HTML File ####
Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('MM/dd/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>Lync Synthetic Transactions Report</title>'
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
add-content $fileName  "<tr bgcolor='#003366'>"
add-content $fileName  "<td colspan='5' height='25' align='center'>"
add-content $fileName  "<font face='Segoe UI' color='#FFFFFF' size='4'><strong>Lync Synthetic Transactions Report - $date</strong></font>"
add-content $fileName  "</td>"
add-content $fileName  "</tr>"
add-content $fileName  "</table>"
}

Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr bgcolor=#0099CC>"
Add-Content $fileName "<td width='40%' align='center'><font color=#FFFFFF>File Location</font></td>"
Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Number of tests</font></td>"
Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Number of failed tests</font></td>"
Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Success Rate</font></td>"
Add-Content $fileName "<td width='15%' align='center'><font color=#FFFFFF>Failure Rate</font></td>"
Add-Content $fileName "</tr>"
}

Function writeServiceInfo
{
param($fileName,$WTestName,$WNumOfRows,$WSucessRate,$WFailureRate,$WFailedTests)

$PerSucess = "{0:P}" -f $WSucessRate
$PerFailure = "{0:P}" -f $WFailureRate

Add-Content $fileName "<tr>"
Add-Content $fileName "<td align=center>$CSVFolder\$WTestName</td>"
Add-Content $fileName "<td align=center>$WNumOfRows</td>"
Add-Content $fileName "<td align=center>$WFailedTests</td>"
Add-Content $fileName "<td align=center>$PerSucess</td>"

 
 if ($WFailureRate -gt $PerFailureThresholdError)
 {
    Add-Content $fileName "<td align=center bgcolor='#FF0000'>$PerFailure</td>"
 }
 elseif ($WFailureRate -gt $PerFailureThresholdWarning)
 {
    Add-Content $fileName "<td align=center bgcolor='#FFA500'>$PerFailure</td>"
 }
 else
 {
    Add-Content $fileName "<td align=center bgcolor='#00FF00'>$PerFailure</td>"
 }
 Add-Content $fileName "</tr>"
}

Function writeHtmlFooter
{
param($fileName)
Add-Content $fileName "</body>"
Add-Content $fileName "</html>"
}

Function sendEmail
{ param($from,$to,$subject,$smtphost,$htmlFileName)
$body = Get-Content $htmlFileName
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
$msg.isBodyhtml = $true
$smtp.send($msg)
}

Function AddSyntheticTransacation
{ param($TestName)

    $CSVResultSuccess = Import-Csv -Path "$LogsFolder\$CurrentDate\$TestName" | Where {$_.Result -eq "Success"}
    $CSVResultFailure = Import-Csv -Path "$LogsFolder\$CurrentDate\$TestName" | Where {$_.Result -eq "Failure"}
    $CountSuccess = ($CSVResultSuccess | Measure-Object).Count
    $CountFailures = ($CSVResultFailure | Measure-Object).Count
    $CSVNumOfRows = $CountSuccess + $CountFailures
    $CSVSuccessRate = $CountSuccess/$CSVNumOfRows
    $CSVFailureRate = $CountFailures/$CSVNumOfRows
    

    #writeTableHeader $EmailFileName
    Add-Content $EmailFileName "<table width='100%'><tbody>"
    Add-Content $EmailFileName "<tr bgcolor='#0099CC'>"
    Add-Content $EmailFileName "<td width='100%' align='center' colSpan=5><font face='segoe ui' color='#FFFFFF' size='2'><strong>$TestName</strong></font></td>"
    Add-Content $EmailFileName "</tr>"
    writeTableHeader $EmailFileName
    writeServiceInfo $EmailFileName $TestName $CSVNumOfRows $CSVSuccessRate $CSVFailureRate $CountFailures
    Add-Content $EmailFileName "</table>"

}

writeHtmlHeader $EmailFileName

AddSyntheticTransacation $FileName_Test_CsAddressbookservice
AddSyntheticTransacation $FileName_Test_CsAddressBookWebQuery
AddSyntheticTransacation $FileName_Test_CsAVConference
AddSyntheticTransacation $FileName_Test_CsGroupIM
AddSyntheticTransacation $FileName_Test_CsIM
AddSyntheticTransacation $FileName_Test_CsP2PAV
AddSyntheticTransacation $FileName_Test_CsPresence
AddSyntheticTransacation $FileName_Test_CsRegistration
#AddSyntheticTransacation $FileName_Test_CsPstnPeerToPeerCall
AddSyntheticTransacation $FileName_Test_CsAVEdgeConnectivity
AddSyntheticTransacation $FileName_Test_CsDataConference
AddSyntheticTransacation $FileName_Test_CsExumConnectivity
AddSyntheticTransacation $FileName_Test_CsGroupIMJoinLauncher
AddSyntheticTransacation $FileName_Test_CsMCXP2PIM
AddSyntheticTransacation $FileName_Test_CsPstnOutboundCall

writeHtmlFooter $EmailFileName

### Configuring Email Parameters
$date2 = ( get-date ).ToString('MM/dd/yyyy')
sendEmail SyntheticTransactions@domain.com "bacharg@domain.com" "Synthetic Transactions Report - $date2" smtp-gw.domain.com $EmailFileName

#Write-Host "`n`nThe File was generated at the following location: $EmailFileName `n`nOpenning file..." -ForegroundColor Cyan
#Invoke-Item $EmailFileName