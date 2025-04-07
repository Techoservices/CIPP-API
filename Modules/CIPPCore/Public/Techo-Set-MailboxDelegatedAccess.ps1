   <#
.SYNOPSIS
    Grants delegated mailbox access for a specified user on all UserMailbox objects in a tenant.
.DESCRIPTION
    Techo-Set-MailboxDelegatedAccess retrieves all mailboxes of type "UserMailbox" (excluding any with "admin" in the alias)
    and, for each mailbox, checks if the specified AccessUser already has the required access rights.
    If delegated access is not already set, the function applies it using New-ExoRequest.
    
    The following parameters are provided via the CIPP GUI:
      - AccessUser      : Email address of the user to be granted access.
      - Automap       : Boolean value for automapping (true or false).
      - TenantFilter  : The tenant ID/domain for which to run the command.
      - AccessRights  : Array of access rights to assign (e.g., "FullAccess").
#>

function Techo-Set-MailboxDelegatedAccess {
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
        # Define the filter to retrieve user mailboxes, excluding those with "admin" in the alias
        $filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -notlike '*admin*')"
        Write-LogMessage -headers $Headers -API $APIName -message "Fetching mailboxes with filter: $filter" -Sev 'Info' -tenant $TenantFilter
        
        # Retrieve all user mailboxes matching the filter
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ ResultSize = 'unlimited'; Filter = $filter }
        
        if (($mailboxes -eq $null) -or ($mailboxes.Count -eq 0)) {
            Write-LogMessage -headers $Headers -API $APIName -message "No matching mailboxes found." -Sev 'Info' -tenant $TenantFilter
            return "No matching mailboxes found."
        }
        
        # Initialize counters for reporting
        $totalMailboxes = $mailboxes.Count
        $successCount = 0
        $failureCount = 0
        $skippedCount = 0
        
        # Process each mailbox
        foreach ($mailbox in $mailboxes) {
            # Determine the mailbox identity to use (adjust property as needed; here we assume UserPrincipalName)
            $mailboxIdentity = $mailbox.UserPrincipalName
            
            # Retrieve existing mailbox permissions for this mailbox
            $existingPerms = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-MailboxPermission' -cmdParams @{ Identity = $mailboxIdentity }
            $alreadyHasAccess = $false
            if ($existingPerms) {
                foreach ($perm in $existingPerms) {
                    # Check if the AccessUser already exists in permissions and has the required access rights.
                    if (($perm.User -eq $AccessUser) -and ($perm.AccessRights -contains $AccessRights[0])) {
                        $alreadyHasAccess = $true
                        break
                    }
                }
            }
            
            if ($alreadyHasAccess) {
                $skippedCount++
                continue
            }
            
            # Attempt to add mailbox permission
            try {
                $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-MailboxPermission' -cmdParams @{
                    Identity        = $mailboxIdentity
                    User            = $AccessUser
                    AccessRights    = $AccessRights
                    InheritanceType = 'all'
                    Automapping     = $Automap
                }
                $successCount++
                Write-LogMessage -headers $Headers -API $APIName -message "Granted $AccessRights to $AccessUser on $mailboxIdentity with automapping: $Automap" -Sev 'Info' -tenant $TenantFilter
            } catch {
                $failureCount++
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -message "Could not add permissions for $AccessUser on $mailboxIdentity. Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
            }
        }
        
        # If no mailbox required processing, exit early.
        if ($successCount -eq 0 -and $skippedCount -eq $totalMailboxes) {
            $message = "All $totalMailboxes mailboxes already have delegated access for $AccessUser. Nothing to process."
            Write-LogMessage -headers $Headers -API $APIName -message $message -Sev 'Info' -tenant $TenantFilter
            return $message
        }
        
        $message = "Processed $totalMailboxes mailboxes: $successCount successes, $failureCount failures, $skippedCount already set."
        Write-LogMessage -headers $Headers -API $APIName -message $message -Sev 'Info' -tenant $TenantFilter
        return $message
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
        return "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)"
    }
}
