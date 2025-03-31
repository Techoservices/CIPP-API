function Set-TECHOMailboxAccess {
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
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ResultSize = 'unlimited'; Filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -ne 'Admin')"}
        
        foreach ($mailbox in $mailboxes) {
            try {
                $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-MailboxPermission' -cmdParams @{
                    Identity = $mailbox.UserPrincipalName
                    User = $AccessUser
                    AccessRights = $AccessRights
                    InheritanceType = 'all'
                    Automapping = $Automap
                }
                Write-LogMessage -headers $Headers -API $APIName -message "Granted $AccessRights to $AccessUser on $($mailbox.UserPrincipalName) with automapping: $Automap" -Sev 'Info' -tenant $TenantFilter
            } catch {
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -message "Could not add permissions for $AccessUser on $($mailbox.UserPrincipalName). Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
            }
        }
        return "Added $AccessUser to all matching mailboxes with automapping: $Automap, with permissions: $($AccessRights -join ', ')"
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
        return "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)"
    }
}