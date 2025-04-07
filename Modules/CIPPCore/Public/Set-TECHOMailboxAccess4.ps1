function Set-TECHOMailboxAccess4 {
    [CmdletBinding()]
    param (
        $AccessUser,
        [bool]$Automap,
        $TenantFilter,
        $APIName = 'Manage Shared Mailbox Access',
        $Headers,
        [array]$AccessRights
    )

    try {
        $filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -notlike '*admin*')"
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ResultSize = 'unlimited'; Filter = $filter}
        
        $totalMailboxes = $mailboxes.Count
        if ($totalMailboxes -eq 0) {
            Write-LogMessage -headers $Headers -API $APIName -message "No matching mailboxes found." -Sev 'Info' -tenant $TenantFilter
            return "No matching mailboxes found."
        }
        
        $successCount = 0
        $failureCount = 0
        
        foreach ($mailbox in $mailboxes) {
            try {
                $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-MailboxPermission' -cmdParams @{
                    Identity = $mailbox.UserPrincipalName
                    User = $AccessUser
                    AccessRights = $AccessRights
                    InheritanceType = 'all'
                    Automapping = $Automap
                }
                $successCount++
                Write-LogMessage -headers $Headers -API $APIName -message "Granted $AccessRights to $AccessUser on $($mailbox.UserPrincipalName) with automapping: $Automap" -Sev 'Info' -tenant $TenantFilter
            } catch {
                $failureCount++
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -message "Could not add permissions for $AccessUser on $($mailbox.UserPrincipalName). Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
            }
        }
        
        $message = "Processed $totalMailboxes mailboxes: $successCount successes, $failureCount failures."
        Write-LogMessage -headers $Headers -API $APIName -message $message -Sev 'Info' -tenant $TenantFilter
        return $message
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
        return "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)"
    }
}