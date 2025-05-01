function Hello-Techo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessUser,
        [Parameter(Mandatory = $false)]
        [bool]$Automap = $false,
        [Parameter(Mandatory = $true)]
        [array]$AccessRights,
        $TenantFilter,
        $APIName = 'Set Delegated Mailbox Access',
        $Headers
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

        # Retrieve mailboxes using New-ExoRequest
        $mailboxes = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ 
            ResultSize = 'unlimited'
            Filter = $filter 
        }
        
        $mailboxCount = if ($mailboxes) { $mailboxes.Count } else { 0 }
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Retrieved $mailboxCount mailboxes" -Sev Info
        
        if ($null -eq $mailboxes -or $mailboxCount -eq 0) {
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "No mailboxes matched the filter criteria" -Sev Info
            return "No mailboxes matched the filter criteria."
        }
        
        # Process each mailbox
        foreach ($mailbox in $mailboxes) {
            # Check for timeout
            if ((Get-Date) - $startTime -gt $timeout) {
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Timeout reached after 10 minutes. Processed: $($results.Success) added, $($results.Failed) failed, $($results.Skipped) skipped." -Sev Warning
                break
            }
            
            $mailboxIdentity = $mailbox.UserPrincipalName
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Processing mailbox: $mailboxIdentity" -Sev Debug
            
            try {
                # Check existing permissions
                $existingPerm = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-MailboxPermission' -cmdParams @{ 
                    Identity = $mailboxIdentity
                    User = $AccessUser 
                }
                
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
                    $null = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Add-MailboxPermission' -cmdParams @{
                        Identity = $mailboxIdentity
                        User = $AccessUser
                        AccessRights = $AccessRights
                        InheritanceType = 'all'
                        Automapping = $Automap
                    }
                    
                    $results.Success++
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Added $($AccessRights -join ', ') access for $AccessUser to $mailboxIdentity (AutoMapping: $Automap)" -Sev Info
                }
            }
            catch {
                $results.Failed++
                $errorMessage = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Error granting access to $mailboxIdentity: $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
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