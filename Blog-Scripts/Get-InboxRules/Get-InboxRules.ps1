$mailboxes = get-mailbox -RecipientTypeDetails Usermailbox -ResultSize Unlimited
$counter = 1
foreach ($mailbox in $mailboxes)
{
    Write-Host "Working on User #$counter out of #$($mailboxes.count) - $mailbox.UserPrincipalName" -ForegroundColor White -BackgroundColor DarkCyan
    $counter = $counter + 1
    $InbouxRules = Get-InboxRule -Mailbox $mailbox.userPrincipalName
    if ($InbouxRules)
        {
            Write-Host "User $($Mailbox.UserPrincipalName) have rules, exporting to CSV"
            $InbouxRules | Export-Csv ".\$($Mailbox.Alias)-InbouxRules-$(Get-Date -Format yyyy-MM-dd__HH-mm).csv" -NoTypeInformation -Force   
        }
}
