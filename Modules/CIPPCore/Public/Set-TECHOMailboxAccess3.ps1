function Set-TECHOMailboxAccess {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$AccessUser,
        [Parameter(Mandatory)]
        [bool]$Automap,
        [Parameter(Mandatory)]
        [string]$TenantFilter,
        [string]$APIName = 'Manage Shared Mailbox Access',
        $Headers,
        [Parameter(Mandatory)]
        [string[]]$AccessRights
    )

    try {
        # Get all user mailboxes excluding admin accounts
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-Mailbox" -cmdParams @{
            ResultSize = "Unlimited"
            Filter = "RecipientTypeDetails -eq 'UserMailbox' -and -not (Alias -like '*admin*')"
        }

        $successCount = 0
        $totalCount = ($mailboxes | Measure-Object).Count

        foreach ($mailbox in $mailboxes) {
            try {
                $null = New-ExoRequest -tenantid $TenantFilter -cmdlet "Add-MailboxPermission" -cmdParams @{
                    Identity = $mailbox.UserPrincipalName
                    User = $AccessUser
                    AccessRights = $AccessRights
                    InheritanceType = "All"
                    AutoMapping = $Automap
                }
                $successCount++
            }
            catch {
                Write-LogMessage -headers $Headers -API $APIName -message "Failed to add permissions for mailbox $($mailbox.UserPrincipalName): $($_.Exception.Message)" -Sev "Warning" -tenant $TenantFilter
                continue
            }
        }

        $automapStatus = if ($Automap) { "with" } else { "without" }
        $message = "Successfully processed $successCount out of $totalCount mailboxes. Added $AccessUser $automapStatus automapping, permissions: $($AccessRights -join ',')"
        Write-LogMessage -headers $Headers -API $APIName -message $message -Sev "Info" -tenant $TenantFilter
        return $message
    }
    catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not process mailbox permissions for $AccessUser. Error: $($ErrorMessage.NormalizedError)" -Sev "Error" -tenant $TenantFilter -LogData $ErrorMessage
        throw "Could not process mailbox permissions. Error: $($ErrorMessage.NormalizedError)"
    }
}