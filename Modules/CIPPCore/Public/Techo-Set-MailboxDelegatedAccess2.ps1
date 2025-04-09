<#
.SYNOPSIS
    Grants delegated mailbox access for a specified user on all UserMailbox objects in a tenant.

.DESCRIPTION
    Grants specified access rights to all user mailboxes (excluding those with 'admin' in the alias) to the specified user. Designed to run daily to include new users.

    The following parameters are provided via the CIPP GUI:
      - AccessUser      : Email address of the user to be granted access.
      - Automap         : Boolean value for automapping (true or false).
      - TenantFilter    : The tenant ID/domain for which to run the command.
      - AccessRights    : Array of access rights to assign (e.g., "FullAccess").
#>

function Techo-Set-MailboxDelegatedAccess2 {
    [CmdletBinding()]
    param (
        $AccessUser,
        [bool]$Automap,
        $TenantFilter,
        $APIName = 'Manage Shared Mailbox Access',
        $Headers,
        [array]$AccessRights
    )

    $startTime = Get-Date
    $timeout = New-TimeSpan -Minutes 10

    try {
        # Define the filter to retrieve user mailboxes, excluding those with "admin" in the alias
        $filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -notlike '*admin*')"
        Write-LogMessage -headers $Headers -API $APIName -message "Fetching mailboxes with filter: $filter at $(Get-Date)" -Sev 'Info' -tenant $TenantFilter
        
        # Retrieve all user mailboxes matching the filter
        $mailboxStart = Get-Date
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ ResultSize = 'unlimited'; Filter = $filter }
        $mailboxEnd = Get-Date
        Write-LogMessage -headers $Headers -API $APIName -message "Get-Mailbox completed in $($mailboxEnd - $mailboxStart). Count: $($mailboxes.Count)" -Sev 'Info' -tenant $TenantFilter
        
        if (-not $mailboxes) {
            Write-LogMessage -headers $Headers -API $APIName -message "No matching mailboxes found at $(Get-Date)" -Sev 'Info' -tenant $TenantFilter
            return "No matching mailboxes found."
        }
        
        # Initialize counters for reporting
        $totalMailboxes = $mailboxes.Count
        $successCount = 0
        $failureCount = 0
        $skippedCount = 0
        
        # Process each mailbox
        foreach ($mailbox in $mailboxes) {
            # Check for timeout
            if ((Get-Date) - $startTime -gt $timeout) {
                Write-LogMessage -headers $Headers -API $APIName -message "Timeout of 10 minutes reached at $(Get-Date). Stopping processing. Processed $successCount successes, $failureCount failures, $skippedCount skipped." -Sev 'Warn' -tenant $TenantFilter
                break
            }
            
            $mailboxIdentity = $mailbox.UserPrincipalName  # Adjust if necessary
            
            # Retrieve existing mailbox permissions for this user on this mailbox
            $permStart = Get-Date
            $existingPerm = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-MailboxPermission' -cmdParams @{ Identity = $mailboxIdentity; User = $AccessUser }
            $permEnd = Get-Date
            Write-LogMessage -headers $Headers -API $APIName -message "Get-MailboxPermission for $mailboxIdentity took $($permEnd - $permStart)" -Sev 'Debug' -tenant $TenantFilter
            
            # Check if the user already has all required access rights
            $hasAllRights = $true
            if ($existingPerm) {
                foreach ($right in $AccessRights) {
                    if ($right -not in $existingPerm.AccessRights) {
                        $hasAllRights = $false
                        break
                    }
                }
            } else {
                $hasAllRights = $false
            }
            
            if ($hasAllRights) {
                $skippedCount++
                Write-LogMessage -headers $Headers -API $APIName -message "Skipped $mailboxIdentity for $AccessUser - already has all rights at $(Get-Date)" -Sev 'Info' -tenant $TenantFilter
                continue
            }
            
            # Attempt to add mailbox permission
            try {
                $addStart = Get-Date
                $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-MailboxPermission' -cmdParams @{
                    Identity        = $mailboxIdentity
                    User            = $AccessUser
                    AccessRights    = $AccessRights
                    InheritanceType = 'all'
                    Automapping     = $Automap
                }
                $addEnd = Get-Date
                Write-LogMessage -headers $Headers -API $APIName -message "Add-MailboxPermission for $mailboxIdentity took $($addEnd - $addStart)" -Sev 'Debug' -tenant $TenantFilter
                $successCount++
                Write-LogMessage -headers $Headers -API $APIName -message "Granted $AccessRights to $AccessUser on $mailboxIdentity with automapping: $Automap at $(Get-Date)" -Sev 'Info' -tenant $TenantFilter
            } catch {
                $failureCount++
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -message "Could not add permissions for $AccessUser on $mailboxIdentity at $(Get-Date). Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
            }
        }
        
        # Summary of processing results
        $elapsedTime = (Get-Date) - $startTime
        $message = "Processed $totalMailboxes mailboxes in $elapsedTime: $successCount successes, $failureCount failures, $skippedCount already set."
        Write-LogMessage -headers $Headers -API $APIName -message $message -Sev 'Info' -tenant $TenantFilter
        return $message
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not retrieve mailboxes at $(Get-Date). Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
        return "Could not retrieve mailboxes. Error: $($ErrorMessage.NormalizedError)"
    }
}