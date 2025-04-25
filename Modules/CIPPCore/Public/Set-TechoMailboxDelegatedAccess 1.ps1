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

function Set-TechoMailboxDelegatedAccess {
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
        # Set timeout parameters
        $startTime = Get-Date
        $timeoutMinutes = 10
        $script:timedOut = $false

        # Register timeout event
        $job = Register-ObjectEvent -InputObject ([System.Timers.Timer]::new(($timeoutMinutes * 60 * 1000))) -EventName Elapsed -Action {
            $script:timedOut = $true
        } -MaxTriggerCount 1
        $job.Enabled = $true

        # Define the filter to retrieve user mailboxes
        $filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -notlike '*admin*')"
        Write-LogMessage -headers $Headers -API $APIName -message "Starting mailbox permission process with timeout of $timeoutMinutes minutes" -Sev 'Info' -tenant $TenantFilter
        
        # Retrieve all user mailboxes matching the filter
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ 
            ResultSize = 'unlimited'
            Filter = $filter 
        }
        
        if (-not $mailboxes) {
            Write-LogMessage -headers $Headers -API $APIName -message "No matching mailboxes found." -Sev 'Info' -tenant $TenantFilter
            return "No matching mailboxes found."
        }
        
        $totalMailboxes = $mailboxes.Count
        $successCount = 0
        $failureCount = 0
        $skippedCount = 0
        
        foreach ($mailbox in $mailboxes) {
            if ($script:timedOut) {
                $timeoutMessage = "Operation timed out after $timeoutMinutes minutes. Processed $($successCount + $skippedCount) of $totalMailboxes mailboxes."
                Write-LogMessage -headers $Headers -API $APIName -message $timeoutMessage -Sev 'Warning' -tenant $TenantFilter
                return $timeoutMessage
            }

            $mailboxIdentity = $mailbox.UserPrincipalName
            Write-LogMessage -headers $Headers -API $APIName -message "Processing mailbox: $mailboxIdentity" -Sev 'Debug' -tenant $TenantFilter
            
            $existingPerms = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-MailboxPermission' -cmdParams @{ 
                Identity = $mailboxIdentity 
            }

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
        
        # Cleanup timer job
        $job.Enabled = $false
        Unregister-Event -SourceIdentifier $job.Name
        Remove-Job -Job $job -Force

        # Return results
        if ($successCount -eq 0 -and $skippedCount -eq $totalMailboxes) {
            $message = "All $totalMailboxes mailboxes already have delegated access for $AccessUser. Nothing to process."
            Write-LogMessage -headers $Headers -API $APIName -message $message -Sev 'Info' -tenant $TenantFilter
            return $message
        }
        
        $message = "Processed $totalMailboxes mailboxes: $successCount successes, $failureCount failures, $skippedCount already set."
        Write-LogMessage -headers $Headers -API $APIName -message $message -Sev 'Info' -tenant $TenantFilter
        return $message
    }
    catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -message "Could not process mailbox permissions. Error: $($ErrorMessage.NormalizedError)" -Sev 'Error' -tenant $TenantFilter -LogData $ErrorMessage
        throw "Could not process mailbox permissions. Error: $($ErrorMessage.NormalizedError)"
    }
    finally {
        # Ensure cleanup of timer job even if script errors out
        if ($job) {
            $job.Enabled = $false
            Unregister-Event -SourceIdentifier $job.Name -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
}
