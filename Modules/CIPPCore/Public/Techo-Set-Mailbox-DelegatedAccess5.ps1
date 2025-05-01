<#
.SYNOPSIS
    Grants delegated mailbox access for a specified user.

.DESCRIPTION
    Grants specified access rights to all user mailboxes (excluding those with 'admin' in the alias) to a user. Runs daily via CIPP GUI.

    Parameters provided via CIPP GUI:
      - AccessUser: Email address of the user to grant access.
      - Automap: Boolean for automapping (true/false).
      - TenantFilter: Tenant ID/domain.
      - AccessRights: Array of access rights (e.g., "FullAccess").
#>
function Techo-Set-Mailbox-DelegatedAccess5 {
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
        # Define filter to exclude admin mailboxes
        $filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -notlike '*admin*')"
        Write-LogMessage -headers $Headers -API $APIName -message "Fetching mailboxes with filter: $filter at $(Get-Date)" -Sev 'Info' -tenant $TenantFilter
        
        # Retrieve mailboxes
        $mailboxStart = Get-Date
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ ResultSize = 'unlimited'; Filter = $filter }
        $mailboxEnd = Get-Date
        Write-LogMessage -headers $Headers -API $APIName -message "Get-Mailbox completed in $($mailboxEnd - $mailboxStart). Count: $($mailboxes.Count)" -Sev 'Info' -tenant $TenantFilter
        
        if (-not $mailboxes) {
            Write-LogMessage -headers $Headers -API $APIName -message "No matching mailboxes found at $(Get-Date)" -Sev 'Info' -tenant $TenantFilter
            return "No matching mailboxes found."
        }
        
        # Initialize counters
        $totalMailboxes = $mailboxes.Count
        $successCount = 0
        $failureCount = 0
        $skippedCount = 0
        
        # Process each mailbox
        foreach ($mailbox in $mailboxes) {
            # Check timeout
            if ((Get-Date) - $startTime -gt $timeout) {
                Write-LogMessage -headers $Headers -API $APIName -message "Timeout of 10 minutes reached at $(Get-Date). Processed $successCount successes, $failureCount failures, $skippedCount skipped." -Sev 'Warn' -tenant $TenantFilter
                break
            }
            
            $mailboxIdentity = $mailbox.UserPrincipalName
            
            # Check existing permissions for the user
            $permStart = Get-Date
            $existingPerm = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-MailboxPermission' -cmdParams @{ Identity = $mailboxIdentity; User = $AccessUser }
            $permEnd = Get-Date
            Write-LogMessage -headers $Headers -API $APIName -message "Get-MailboxPermission for $mailboxIdentity took $($permEnd - $permStart)" -Sev 'Debug' -tenant $TenantFilter
            
            # Verify all required access rights
            $hasAllRights = $true
            if ($existingPerm) {
                foreach ($right in $AccessRights) {
                    if ($right -notin $existingPerm.AccessRights) {
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
            
            # Add permissions
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
        
        # Summary
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