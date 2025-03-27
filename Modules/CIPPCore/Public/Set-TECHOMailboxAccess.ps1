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
        $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -ne 'Admin')} | Add-MailboxPermission' -cmdParams @{user = $AccessUser; automapping = $Automap; accessRights = $AccessRights; InheritanceType = 'all' }
        
        if ($Automap) {
            Write-LogMessage -headers $Headers -API $APIName -message "Gave $AccessRights permissions to $($AccessUser) on with automapping" -Sev 'Info' -tenant $TenantFilter
            return "Added $($AccessUser) to Shared Mailbox with automapping, with the following permissions: $AccessRights"
        } else {
            Write-LogMessage -headers $Headers -API $APIName -message "Gave $AccessRights permissions to $($AccessUser) on without automapping" -Sev 'Info' -tenant $TenantFilter
            return "Added $($AccessUser) to Shared Mailbox without automapping, with the following permissions: $AccessRights"
        }
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not add mailbox permissions for $($AccessUser) on Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
        return "Could not add shared mailbox permissions for. Error: $($ErrorMessage.NormalizedError)"
    }
}