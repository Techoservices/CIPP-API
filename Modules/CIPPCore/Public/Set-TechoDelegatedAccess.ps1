<#
.SYNOPSIS
    Grants delegated mailbox access for a specified user on all UserMailbox objects in a tenant.

.DESCRIPTION
    Grants specified access rights to all user mailboxes (excluding those with 'admin' in the alias) to a user.
    Designed to run daily via CIPP GUI to include new users.

    Parameters provided via CIPP GUI:
      - AccessUser: Email address of the user to grant access.
      - Automap: Boolean for automapping (true/false).
      - TenantFilter: Tenant ID/domain.
      - AccessRights: Array of access rights (e.g., "FullAccess").
#>

function Set-TechoDelegatedAccess {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$AccessUser,
        
        [Parameter(Mandatory = $false)]
        [bool]$Automap = $false,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf', 'ReadPermission')]
        [array]$AccessRights,
        
        [Parameter(Mandatory = $true)]
        [string]$TenantFilter,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory = $false)]
        [string]$APIName = 'Set Delegated Mailbox Access'
    )

    try {
        $startTime = Get-Date
        $timeout = New-TimeSpan -Minutes 10
        $results = [PSCustomObject]@{
            Success = 0
            Failed = 0
            Skipped = 0
        }

        # Define filter to exclude admin mailboxes
        $filter = "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -notlike '*admin*')"
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Fetching mailboxes with filter: $filter" -Sev Info

        # Retrieve mailboxes using New-ExoRequest with error handling
        $mailboxStart = Get-Date
        try {
            $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ 
                ResultSize = 'unlimited'
                Filter = $filter 
            } -ErrorAction Stop
        }
        catch {
            $errorMessage = Get-CippException -Exception $_
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed to retrieve mailboxes: $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
            throw "Failed to retrieve mailboxes: $($errorMessage.GeneralErrorMessage)"
        }

        $mailboxEnd = Get-Date
        $mailboxCount = if ($mailboxes) { $mailboxes.Count } else { 0 }
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Retrieved $mailboxCount mailboxes in $($mailboxEnd - $mailboxStart)" -Sev Info
        
        if ($null -eq $mailboxes -or $mailboxCount -eq 0) {
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "No mailboxes matched the filter criteria" -Sev Info
            return "No mailboxes matched the filter criteria."
        }
        
        # Process mailboxes in batches of 50
        $batchSize = 50
        $batches = [Math]::Ceiling($mailboxCount / $batchSize)
        
        for ($batch = 0; $batch -lt $batches; $batch++) {
            $startIndex = $batch * $batchSize
            $endIndex = [Math]::Min(($startIndex + $batchSize - 1), ($mailboxCount - 1))
            $currentBatch = $mailboxes[$startIndex..$endIndex]
            
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Processing batch $($batch + 1) of $batches (mailboxes $($startIndex + 1) to $($endIndex + 1))" -Sev Info
            
            # Process each mailbox in the current batch
            foreach ($mailbox in $currentBatch) {
                # Check for timeout
                if ((Get-Date) - $startTime -gt $timeout) {
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Timeout reached after 10 minutes. Processed: $($results.Success) added, $($results.Failed) failed, $($results.Skipped) skipped." -Sev Warning
                    break
                }
                
                $mailboxIdentity = $mailbox.UserPrincipalName
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Processing mailbox: $mailboxIdentity" -Sev Debug
                
                try {
                    # Check existing permissions with retry logic
                    $maxRetries = 3
                    $retryCount = 0
                    $permStart = Get-Date
                    
                    do {
                        try {
                            $existingPerm = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-MailboxPermission' -cmdParams @{ 
                                Identity = $mailboxIdentity
                                User = $AccessUser 
                            } -ErrorAction Stop
                            break
                        }
                        catch {
                            $retryCount++
                            if ($retryCount -eq $maxRetries) { throw }
                            Start-Sleep -Seconds (2 * $retryCount)
                        }
                    } while ($retryCount -lt $maxRetries)
                    
                    $permEnd = Get-Date
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Get-MailboxPermission for $mailboxIdentity took $($permEnd - $permStart)" -Sev Debug
                    
                    # Check if all requested rights are already granted
                    $needsUpdate = $true
                    if ($null -ne $existingPerm) {
                        $hasAllRights = $true
                        
                        # Handle case where existingPerm could be a single object or an array
                        $currentRights = @($existingPerm | ForEach-Object { $_.AccessRights }) | ForEach-Object { $_ }
                        
                        foreach ($right in $AccessRights) {
                            if ($currentRights -notcontains $right) {
                                $hasAllRights = $false
                                break
                            }
                        }
                        
                        if ($hasAllRights) {
                            $needsUpdate = $false
                            $results.Skipped++
                            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Skipping $mailboxIdentity - $AccessUser already has required permissions" -Sev Info
                        }
                    }
                    
                    # Add permissions if needed
                    if ($needsUpdate) {
                        $addStart = Get-Date
                        
                        # Handle SendAs permission separately
                        if ($AccessRights -contains 'SendAs') {
                            $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-RecipientPermission' -cmdParams @{
                                Identity = $mailboxIdentity
                                Trustee = $AccessUser
                                AccessRights = 'SendAs'
                            } -ErrorAction Stop
                        }
                        
                        # Handle other permissions using Add-MailboxPermission
                        $otherRights = $AccessRights | Where-Object { $_ -ne 'SendAs' }
                        if ($otherRights) {
                            $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-MailboxPermission' -cmdParams @{
                                Identity = $mailboxIdentity
                                User = $AccessUser
                                AccessRights = $otherRights
                                InheritanceType = 'all'
                                Automapping = $Automap
                            } -ErrorAction Stop
                        }
                        
                        $addEnd = Get-Date
                        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Added permissions for $mailboxIdentity took $($addEnd - $addStart)" -Sev Debug
                        
                        $results.Success++
                        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Added $($AccessRights -join ', ') access for $AccessUser to $mailboxIdentity (AutoMapping: $Automap)" -Sev Info
                    }
                }
                catch {
                    $results.Failed++
                    $errorMessage = Get-CippException -Exception $_
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Error granting access to $mailboxIdentity - $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
                }
            }
        }
        
        # Generate summary
        $totalTime = (Get-Date) - $startTime
        $formattedTime = "{0:mm}m {0:ss}s" -f $totalTime
        $summary = "Completed in $formattedTime. Processed $mailboxCount mailboxes: $($results.Success) added, $($results.Failed) failed, $($results.Skipped) skipped."
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message $summary -Sev Info
        
        return $summary
    }
    catch {
        $errorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Script execution failed: $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
        return "Script execution failed: $($errorMessage.GeneralErrorMessage)"
    }
} 
