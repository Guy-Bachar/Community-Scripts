#------------------------------------------------------------------------------------------
# Modules Connection Status
#------------------------------------------------------------------------------------------
CLS
Write-Host "[Phase 1] - Checking Azure AD Connect and Module Status" -ForegroundColor Magenta
#Connect-AzureAD
#Connect-MsolService
#Connect-ExchangeOnlinev2

#------------------------------------------------------------------------------------------
# Sync Status
#------------------------------------------------------------------------------------------
Write-Host "[Phase 2] - Tenant Directory Sync" -ForegroundColor Magenta
$CurrentADConnectState = Get-MsolCompanyInformation
$TimerSeconds = 60
if ($CurrentADConnectState.DirectorySynchronizationEnabled -eq $true)
{
    Write-Host "`t[Phase 2] - Directory Syncronization is Enabled and Needs to be Stopped First" -ForegroundColor Yellow
    Write-Host "`t[Phase 2] - Type: Set-MsolDirSyncEnabled -EnableDirSync $false" -ForegroundColor Yellow
    #Set-MsolDomainAuthentication -DomainName <domain> -Authentication Managed
}
else
{
    Write-Host "`t[Phase 2] - Directory Sync is Disabled, Proceeding..." -ForegroundColor Green
    Write-Host "`t[Phase 2] - Tenant Sync Status is: $($(Get-MSOLCompanyInformation).DirectorySynchronizationStatus)" -ForegroundColor Cyan

    while($(Get-MSOLCompanyInformation).DirectorySynchronizationStatus -ne "Disabled") #Enabled
     {
        Write-Host "`t[Phase 2] - Tenant Status is not yet disabled, Checking again in 5 minutues, Current Status: $($(Get-MSOLCompanyInformation).DirectorySynchronizationStatus)" -ForegroundColor Yellow
        Write-host "`t[Phase 2] - Current Amount of DirectorySync Enabled Users" -ForegroundColor Yellow
        Get-AzureADUser -All $true | Group-Object -Property:DirSyncEnabled | ft -AutoSize
        #Get-AzureADUser | Where {$_.DirSyncEnabled -eq $true} | Select -Property DisplayName,UserPrincipalName,mail,DirSyncEnabled,LastDirSyncTime | ft -auto
        Write-host "`t[Phase 2] - Current Amount of DirectorySync Enabled Groups" -ForegroundColor Yellow
        Get-AzureADGroup -All $true | Group-Object -Property:DirSyncEnabled | ft -AutoSize
        #Get-AzureADGroup | Where {$_.DirSyncEnabled -eq $true} | Select -Property DisplayName,UserPrincipalName,mail,DirSyncEnabled,LastDirSyncTime | ft -auto
        Write-Host "`t[Phase 2] - Checking again in $$TimerSeconds"
        Start-Sleep -Seconds $TimerSeconds
     }
}


#------------------------------------------------------------------------------------------
# Domain Removal Process
#------------------------------------------------------------------------------------------
Write-Host "[Phase 3] - Objects Query" -ForegroundColor Magenta
Write-Host "`t[Phase 3] - Quering Mailboxes" -ForegroundColor Green -NoNewline
$Mailboxes = Get-EXOMailbox -ResultSize Unlimited | Select-Object
Write-Host ", Found: $($Mailboxes.Count)" -ForegroundColor Green

Write-Host "`t[Phase 3] - Quering Unified Groups" -ForegroundColor Green -NoNewline
$UnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited
Write-Host ", Found: $($UnifiedGroups.Count)" -ForegroundColor Green

Write-Host "`t[Phase 3] - Quering Recipients" -ForegroundColor Green -NoNewline
$Recipients = Get-EXORecipient -ResultSize Unlimited
Write-Host ", Found: $($Recipients.Count)" -ForegroundColor Green

Write-Host "`t[Phase 3] - Quering Contacts" -ForegroundColor Green -NoNewline
$Contacts = Get-MailContact -ResultSize Unlimited
Write-Host ", Found: $($Contacts.Count)" -ForegroundColor Green

Write-Host "`t[Phase 3] - Quering DistriubtionGroups" -ForegroundColor Green -NoNewline
$DLs = Get-DistributionGroup -ResultSize Unlimited
Write-Host ", Found: $($DsL.Count)" -ForegroundColor Green

Write-Host "`t[Phase 3] - Quering DistriubtionGroups" -ForegroundColor Green -NoNewline
$MailUsers = Get-MailUser -ResultSize Unlimited
Write-Host ", Found: $($MailUsers.Count)" -ForegroundColor Green

#------------------------------------------------------------------------------------------
# Domains:
#   Domain1 - domain.com
#------------------------------------------------------------------------------------------
Write-Host "[Phase 4] - Domain Filtering, Identification and Default Domain Set" -ForegroundColor Magenta
$DomainToRemove = "Domain.com"
$TenantDomain = (Get-MsolDomain | Where { $_.Name -like "*.onmicrosoft.com" -and $_.Name -notlike "*.mail.onmicrosoft.com" }).Name

#------------------------------------------------------------------------------------------
# Changing Default Domain to OnMicrosoft
#------------------------------------------------------------------------------------------
$DefaultDomain = Get-MsolDomain | where {$_.IsDefault -eq $true}
if ($DefaultDomain.name -eq $TenantDomain) {Write-Host "`t[Phase 4] - Default doamin matching $TenantDomain" -ForegroundColor Green}
else {Write-Host "`t[Phase 4] - Current Default Domain is $($DefaultDomain.Name) , Changing Default Domain to $TenantDomain" -ForegroundColor Yellow -NoNewline ; Set-MsolDomain -Name $TenantDomain -IsDefault ; Write-Host "[Phase 4] - Done. Allowing 5 Minutes for Replication." -ForegroundColor Green;Start-Sleep -Seconds 300}


#------------------------------------------------------------------------------------------
# UPN Change
#------------------------------------------------------------------------------------------
$MailboxCounter = 1
$UPNChanges = @()
Write-Host "[Phase 5] - UPN Change" -ForegroundColor Magenta

foreach ($mailbox in $Mailboxes)
    {
    Write-Host "`t[Phase 5] - Mailboxes - [$MailboxCounter out of $($Mailboxes.count)] - Working on $($mailbox.DisplayName)" -ForegroundColor Green -NoNewline; Write-Host " ($($mailbox.RecipientTypeDetails))" -ForegroundColor Yellow ; $MailboxCounter += 1
    $tempproxyaddress = $null
    $tempproxyaddress = $mailbox.EmailAddresses | Where { $_ -like 'smtp*' }
    if ($tempproxyaddress.count -gt 0) {
        $onMicrosoftAddress = $null
        $onMicrosoftAddress = $tempproxyaddress | where { $_ -like "*$TenantDomain" }
        try
        {
        Set-MsolUserPrincipalName -UserPrincipalName $mailbox.UserPrincipalName -NewUserPrincipalName $onMicrosoftAddress.Split(":")[1]
        $UPNChanges += New-Object PSObject -property @{
                    DisplayName          = $mailbox.DisplayName
                    UserPrincipalName    = $mailbox.UserPrincipalName
                    Alias                = $mailbox.Alias
                    EmailRemoved         = $proxyaddress
                    SMTPEmailDomainName  = $proxyaddress.Split("@")[1]
                    RecipientTypeDetails = $mailbox.RecipientTypeDetails
                    PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                    IsDirSynced          = $mailbox.IsDirSynced
                    TypeOfAddress        = if ($proxyaddress -clike 'SMTP:*') { "Primary" } else { "Secondary" }
                }
        }
        catch
        {
         Write-Error "There was an error change UPN"   
        }
        }
    }

#Export
$UPNChanges | Export-Csv -NoTypeInformation ".\$(Get-Date -Format 'yyyy_MM_dd-HH_mm')_MailboxesEmailRemoval.csv"


#------------------------------------------------------------------------------------------
# Mailboxes Proxy Address Removal
#------------------------------------------------------------------------------------------
$MailboxCounter = 1
$EmailAddressesRemoved = @()
Write-Host "[Phase 6] - Mailboxes Removal" -ForegroundColor Magenta

foreach ($mailbox in $Mailboxes) {
    Write-Host "`t[Phase 6] - Mailboxes - [$MailboxCounter out of $($Mailboxes.count)] - Working on $($mailbox.DisplayName)" -ForegroundColor Green -NoNewline; Write-Host " ($($mailbox.RecipientTypeDetails))" -ForegroundColor Yellow ; $MailboxCounter += 1
    $tempproxyaddress = $null
    $tempproxyaddress = $mailbox.EmailAddresses | Where { $_ -like 'smtp*' }
    if ($tempproxyaddress.count -gt 0) {
        $onMicrosoftAddress = $null
        $onMicrosoftAddress = $tempproxyaddress | where { $_ -like "*$TenantDomain" }
        Set-Mailbox -Identity $mailbox.Id -WindowsEmailAddress $onMicrosoftAddress.Split(":")[1]

        foreach ($proxyaddress in $tempproxyaddress) {
            if ($proxyaddress -like "*$DomainToRemove") {

                Write-Host "`t`t $($DomainToRemove) email domain was found - $proxyaddress" -ForegroundColor Gray
                Set-Mailbox -Identity $mailbox.Id -EmailAddresses @{Remove = $proxyaddress } -Verbose #-WhatIf

                # Adding User to Removed Users
                $EmailAddressesRemoved += New-Object PSObject -property @{
                    DisplayName          = $mailbox.DisplayName
                    UserPrincipalName    = $mailbox.UserPrincipalName
                    Alias                = $mailbox.Alias
                    EmailRemoved         = $proxyaddress
                    SMTPEmailDomainName  = $proxyaddress.Split("@")[1]
                    RecipientTypeDetails = $mailbox.RecipientTypeDetails
                    PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                    IsDirSynced          = $mailbox.IsDirSynced
                    TypeOfAddress        = if ($proxyaddress -clike 'SMTP:*') { "Primary" } else { "Secondary" }
                }
            }
            else {
                Write-Host "`t`t No action taken against the following address- $proxyaddress"
            }
        }
        
    }
    else {
        Write-Warning "`tNo Proxy Addresses were found"
    }
}
#Export
$EmailAddressesRemoved | Export-Csv -NoTypeInformation ".\$(Get-Date -Format 'yyyy_MM_dd-HH_mm')_MailboxesEmailRemoval.csv"


# ------------------------------------------------------------------------------------------
# Unified Groups Proxy Address Removal
# ------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------
# Examples of Removal based on the results you are getting in the Domain Removal Page:
#Set-UnifiedGroup -Identity 'Outlook Customer Manager' -Alias @{remove="smtp:AllSalesTeam-c036c9b0-d80c-423b-afed-8642ce2d6076@mosquitosquad.co"}
#Set-UnifiedGroup -identity MyGroupName -EmailAddresses @{Remove="smtp:myGroup@domain_we_want_to_remove"}
#Set-UnifiedGroup -identity MyGroupName -EmailAddresses @{Remove="myGroup@domain_we_want_to_remove"}
# ------------------------------------------------------------------------------------------
Write-Host "[Phase 7] - Unified Groups Removal" -ForegroundColor Magenta
$UnifiedGroupCounter = 1
$UnifiedGroupEmailAddressesRemoved = @()

foreach ($unifiedgroup in $UnifiedGroups) {
    Write-Host "`t[Phase 7] - UnifiedGroups - [$UnifiedGroupCounter out of $($UnifiedGroups.count)] - Working on $($unifiedgroup.DisplayName)" -ForegroundColor Green -NoNewline; Write-Host " ($($unifiedgroup.RecipientTypeDetails))" -ForegroundColor Yellow ; $UnifiedGroupCounter += 1
    $tempproxyaddress = $null
    $tempproxyaddress = $unifiedgroup.EmailAddresses | Where { $_ -like 'smtp*' }
    if ($tempproxyaddress.count -gt 0) {

        $onMicrosoftAddressUnifiedGroup = $null
        $onMicrosoftAddressUnifiedGroup = $tempproxyaddress | where { $_ -like "*$TenantDomain" }
        if ($unifiedgroup.PrimarySmtpAddress -ne $($onMicrosoftAddressUnifiedGroup.Split(":")[1]))
        {
            Set-UnifiedGroup -Identity $unifiedgroup.Id -PrimarySmtpAddress $onMicrosoftAddressUnifiedGroup.Split(":")[1]
        }

        foreach ($proxyaddress in $tempproxyaddress) {
            if ($proxyaddress -like "*$DomainToRemove") {
                Write-Host "`t`t $($DomainToRemove) email domain was found - $proxyaddress" -ForegroundColor Gray
                Set-UnifiedGroup -Identity $unifiedgroup.Id -EmailAddresses @{Remove = $proxyaddress }

                # Adding User to Removed Users
                $UnifiedGroupEmailAddressesRemoved += New-Object PSObject -property @{
                    DisplayName          = $unifiedgroup.DisplayName
                    #UserPrincipalName = $unifiedgroup.UserPrincipalName
                    Alias                = $unifiedgroup.Alias
                    EmailRemoved         = $proxyaddress
                    SMTPEmailDomainName  = $proxyaddress.Split("@")[1]
                    RecipientTypeDetails = $unifiedgroup.RecipientTypeDetails
                    PrimarySmtpAddress   = $unifiedgroup.PrimarySmtpAddress
                    IsDirSynced          = $unifiedgroup.IsDirSynced
                    TypeOfAddress        = if ($proxyaddress -clike 'SMTP:*') { "Primary" } else { "Secondary" }
                }
            }
            else {
                Write-Host "`t No action taken against the following address- $proxyaddress"
            }
        }
    }
    else {
        Write-Warning "No Proxy Addresses were found"
    }
}
#Export
$UnifiedGroupEmailAddressesRemoved | Export-Csv -NoTypeInformation ".\$(Get-Date -Format 'yyyy_MM_dd-HH_mm')_UnifiedGrouspEmailRemoval.csv"



#------------------------------------------------------------------------------------------
# Contacts Proxy Address Removal
#------------------------------------------------------------------------------------------
Write-Host "[Phase 8] - Contacts Removal" -ForegroundColor Magenta
$ContactsCounter = 1
$ContactsEmailAddressesRemoved = @()

if ($Contacts.Count -gt 0) {

foreach ($contact in $Contacts) {
    Write-Host "`t[Phase 8] - Contacts - [$ContactsCounter out of $($Contacts.count)] - Working on $($contact.DisplayName)" -ForegroundColor Green -NoNewline; Write-Host " ($($contact.RecipientTypeDetails))" -ForegroundColor Yellow ; $ContactsCounter += 1
    $tempproxyaddress = $null
    $tempproxyaddress = $contact.EmailAddresses | Where { $_ -like 'smtp*' }
    if ($tempproxyaddress.count -gt 0) {
        $onMicrosoftAddress = $null
        $onMicrosoftAddress = $tempproxyaddress | where { $_ -like "*$TenantDomain" }
        #Set-Contact -Identity $contact.Id -WindowsEmailAddress $onMicrosoftAddress.Split(":")[1]

        foreach ($proxyaddress in $tempproxyaddress) {
            if ($proxyaddress -like "*$DomainToRemove") {

                Write-Host "`t`t $($DomainToRemove) email domain was found - $proxyaddress" -ForegroundColor Gray
                if ($contact.PrimarySmtpAddress -eq $($proxyaddress.Split(":")[0]))
                {
                    Write-Host "`t`tThe Address is Primary SMTP Address, Skipping"
                }
                else
                {
                Set-Contact -Identity $contact.Id -EmailAddresses @{Remove = $proxyaddress } -WhatIf
                # Adding User to Removed Users
                $ContactsEmailAddressesRemoved += New-Object PSObject -property @{
                    DisplayName          = $contact.DisplayName
                    UserPrincipalName    = $contact.UserPrincipalName
                    Alias                = $contact.Alias
                    EmailRemoved         = $proxyaddress
                    SMTPEmailDomainName  = $proxyaddress.Split("@")[1]
                    RecipientTypeDetails = $contact.RecipientTypeDetails
                    PrimarySmtpAddress   = $contact.PrimarySmtpAddress
                    IsDirSynced          = $contact.IsDirSynced
                    TypeOfAddress        = if ($proxyaddress -clike 'SMTP:*') { "Primary" } else { "Secondary" }
                }
                }

            }
            else {
                Write-Host "`t`t No action taken against the following address- $proxyaddress"
            }
        }
    }
    else {
        Write-Warning "`tNo Proxy Addresses were found"
    }
}}
#Export
$ContactsEmailAddressesRemoved | Export-Csv -NoTypeInformation ".\$(Get-Date -Format 'yyyy_MM_dd-HH_mm')_ContactsEmailRemoval.csv"




# ------------------------------------------------------------------------------------------
# DL Groups Proxy Address Removal
# ------------------------------------------------------------------------------------------
# ------------------------------------------------------------------------------------------
# Examples of Removal based on the results you are getting in the Domain Removal Page:
#Set-DLGroup -Identity 'Outlook Customer Manager' -Alias @{remove="smtp:AllSalesTeam-c036c9b0-d80c-423b-afed-8642ce2d6076@mosquitosquad.co"}
#Set-DLGroup -identity MyGroupName -EmailAddresses @{Remove="smtp:myGroup@domain_we_want_to_remove"}
#Set-DL -identity MyGroupName -EmailAddresses @{Remove="myGroup@domain_we_want_to_remove"}
# ------------------------------------------------------------------------------------------
Write-Host "[Phase 9] - DL Groups Removal" -ForegroundColor Magenta
$DLGroupCounter = 1
$DLGroupEmailAddressesRemoved = @()

foreach ($DL in $DLs) {
    Write-Host "`t[Phase 6] - DLs - [$DLGroupCounter out of $($DLs.count)] - Working on $($DL.DisplayName)" -ForegroundColor Green -NoNewline; Write-Host " ($($DL.RecipientTypeDetails))" -ForegroundColor Yellow ; $DLGroupCounter += 1
    $tempproxyaddress = $null
    $tempproxyaddress = $DL.EmailAddresses | Where { $_ -like 'smtp*' }
    if ($tempproxyaddress.count -gt 0) {

        $onMicrosoftAddressDLGroup = $null
        $onMicrosoftAddressDLGroup = $tempproxyaddress | where { $_ -like "*$TenantDomain" }
        if ($DL.PrimarySmtpAddress -ne $($onMicrosoftAddressDLGroup.Split(":")[1]))
        {
            Set-DistributionGroup -Identity $DL.Id -PrimarySmtpAddress $onMicrosoftAddressDLGroup.Split(":")[1]
        }

        foreach ($proxyaddress in $tempproxyaddress) {
            if ($proxyaddress -like "*$DomainToRemove") {
                Write-Host "`t`t $($DomainToRemove) email domain was found - $proxyaddress" -ForegroundColor Gray
                Set-DistributionGroup -Identity $DL.Id -EmailAddresses @{Remove = $proxyaddress }

                # Adding User to Removed Users
                $DLGroupEmailAddressesRemoved += New-Object PSObject -property @{
                    DisplayName          = $DL.DisplayName
                    #UserPrincipalName = $DL.UserPrincipalName
                    Alias                = $DL.Alias
                    EmailRemoved         = $proxyaddress
                    SMTPEmailDomainName  = $proxyaddress.Split("@")[1]
                    RecipientTypeDetails = $DL.RecipientTypeDetails
                    PrimarySmtpAddress   = $DL.PrimarySmtpAddress
                    IsDirSynced          = $DL.IsDirSynced
                    TypeOfAddress        = if ($proxyaddress -clike 'SMTP:*') { "Primary" } else { "Secondary" }
                }
            }
            else {
                Write-Host "`t No action taken against the following address- $proxyaddress"
            }
        }
    }
    else {
        Write-Warning "No Proxy Addresses were found"
    }
}
#Export
$DLGroupEmailAddressesRemoved | Export-Csv -NoTypeInformation ".\$(Get-Date -Format 'yyyy_MM_dd-HH_mm')_DLGrouspEmailRemoval.csv"





#------------------------------------------------------------------------------------------
# MailUsers Proxy Address Removal
#------------------------------------------------------------------------------------------
Write-Host "[Phase 10] - MailUsers Removal" -ForegroundColor Magenta
$MailUsersCounter = 1
$MailUsersEmailAddressesRemoved = @()

if ($MailUsers.Count -gt 0) {

foreach ($MailUser in $MailUsers) {
    Write-Host "`t[Phase 10] - MailUsers - [$MailUsersCounter out of $($MailUsers.count)] - Working on $($MailUser.DisplayName)" -ForegroundColor Green -NoNewline; Write-Host " ($($MailUser.RecipientTypeDetails))" -ForegroundColor Yellow ; $MailUsersCounter += 1
    $tempproxyaddress = $null
    $tempproxyaddress = $MailUser.EmailAddresses | Where { $_ -like 'smtp*' }
    if ($tempproxyaddress.count -gt 0) {
        $onMicrosoftAddress = $null
        $onMicrosoftAddress = $tempproxyaddress | where { $_ -like "*$TenantDomain" }
        Set-MailUser -Identity $MailUser.Id -WindowsEmailAddress $onMicrosoftAddress.Split(":")[1]

        foreach ($proxyaddress in $tempproxyaddress) {
            if ($proxyaddress -like "*$DomainToRemove") {

                Write-Host "`t`t $($DomainToRemove) email domain was found - $proxyaddress" -ForegroundColor Gray
                if ($MailUser.PrimarySmtpAddress -eq $($proxyaddress.Split(":")[0]))
                {
                    Write-Host "`t`tThe Address is Primary SMTP Address, Skipping"
                }
                else
                {
                Set-MailUser -Identity $MailUser.Id -EmailAddresses @{Remove = $proxyaddress } -WhatIf
                # Adding User to Removed Users
                $MailUsersEmailAddressesRemoved += New-Object PSObject -property @{
                    DisplayName          = $MailUser.DisplayName
                    UserPrincipalName    = $MailUser.UserPrincipalName
                    Alias                = $MailUser.Alias
                    EmailRemoved         = $proxyaddress
                    SMTPEmailDomainName  = $proxyaddress.Split("@")[1]
                    RecipientTypeDetails = $MailUser.RecipientTypeDetails
                    PrimarySmtpAddress   = $MailUser.PrimarySmtpAddress
                    IsDirSynced          = $MailUser.IsDirSynced
                    TypeOfAddress        = if ($proxyaddress -clike 'SMTP:*') { "Primary" } else { "Secondary" }
                }
                }

            }
            else {
                Write-Host "`t`t No action taken against the following address- $proxyaddress"
            }
        }
    }
    else {
        Write-Warning "`tNo Proxy Addresses were found"
    }
}}
#Export
$MailUsersEmailAddressesRemoved | Export-Csv -NoTypeInformation ".\$(Get-Date -Format 'yyyy_MM_dd-HH_mm')_MailUsersEmailRemoval.csv"


