<#
.SYNOPSIS
    Apply/Update transport rules to add a warning disclaimer for potential display name spoofing in Exchange Online.

.DESCRIPTION
    Idempotent executive command for CIPP:
    - Retrieves all unique mailbox display names.
    - Batches them (200 per rule – safe limit).
    - Creates, updates, or removes rules to exactly match current users.
    - Manages the exception placeholder rule (priority 0, disabled by default).
    - Designed for safe repeated runs via CIPP Executive Commands.

.NOTES
    Runs in CIPP context – uses existing app-only EXO connection via New-ExoRequest.
    No parameters required – works on the selected tenant.
#>

function Set-Techo-DisplayNameSpoofing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantFilter,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $false)]
        [string]$APIName = 'Set Display Name Spoofing Protection'
    )

    try {
        # Full original HTML disclaimer (orange bar + red warning text)
        $ruleHtml = @'
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width="100%" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184;mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>
  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'>
    <td style='background:#ff6200;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td>
    <td width="98%" style='width:98.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt;word-wrap:break-word'>
      <div>
        <p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:5.25pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly'>
          <strong><b><i>
            <span style='font-size:15.0pt;font-family:"Segoe UI",sans-serif;color:#fc0303'>
              This message was sent from outside the company by someone with a display name matching a user in your organization. Please do not click any links, or open attachments, unless you recognize the source of this email and know the content is safe.
            </span>
          </i></b></strong>
        </p>
      </div>
    </td>
  </tr>
</table>
'@

        $ruleNamePrefix      = "Add Disclaimer for Internal Display Names"
        $ExceptionHeader     = "X-Positive-IgnoreDisplayName"
        $ExceptionHeaderText = "True"
        $exceptionRuleName   = "Exceptions to Display Name Disclaimer"
        $exceptionComments   = "This rule is for matching emails that will not be checked for Display Name spoofing. Please use appropriate caution when adding rules, and try to make sure excepted emails are appropriately authenticated"
        $batchSize           = 200

        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Starting Display Name Spoofing protection setup" -Sev Info

        # Retrieve unique display names via New-ExoRequest
        $displayNames = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-Mailbox" -cmdParams @{ ResultSize = "Unlimited" } |
                        Select-Object -ExpandProperty DisplayName |
                        Where-Object { $_.Trim() -ne '' } |
                        Sort-Object -Unique

        if ($displayNames.Count -eq 0) {
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "No valid display names found – nothing to do." -Sev Warning
            return "No mailboxes with display names found."
        }

        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Retrieved $($displayNames.Count) unique display names." -Sev Info

        # Create batches
        $batches = @()
        for ($i = 0; $i -lt $displayNames.Count; $i += $batchSize) {
            $end = [Math]::Min($i + $batchSize - 1, $displayNames.Count - 1)
            $batches += ,@($displayNames[$i..$end])
        }

        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Created $($batches.Count) batch(es) of up to $batchSize names." -Sev Info

        # Get existing rules
        $existingRules = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-TransportRule" |
                         Where-Object { $_.Name -like "$ruleNamePrefix *" } |
                         Sort-Object { if ($_.Name -match '\s(\d+)$') { [int]$matches[1] } else { 9999 } }

        $maxIndex = [Math]::Max($batches.Count, $existingRules.Count) - 1

        for ($i = 0; $i -le $maxIndex; $i++) {
            $ruleNumber   = $i + 1
            $ruleName     = "$ruleNamePrefix $ruleNumber"
            $currentBatch = if ($i -lt $batches.Count) { $batches[$i] } else { $null }

            try {
                if ($null -eq $currentBatch -and $i -lt $existingRules.Count) {
                    # Remove excess
                    $rule = $existingRules[$i]
                    New-ExoRequest -tenantid $TenantFilter -cmdlet "Remove-TransportRule" -cmdParams @{ Identity = $rule.Identity; Confirm = $false }
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Removed excess rule: $ruleName" -Sev Info
                }
                elseif ($i -lt $existingRules.Count) {
                    # Update existing
                    $rule = $existingRules[$i]
                    New-ExoRequest -tenantid $TenantFilter -cmdlet "Set-TransportRule" -cmdParams @{
                        Identity             = $rule.Identity
                        Name                 = $ruleName
                        HeaderContainsWords  = $currentBatch
                        Priority             = ($i + 1)
                    }
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Updated rule $ruleName ($($currentBatch.Count) names)" -Sev Info
                }
                else {
                    # Create new
                    $params = @{
                        Name                                = $ruleName
                        Priority                            = ($i + 1)
                        FromScope                           = "NotInOrganization"
                        HeaderContainsMessageHeader         = "From"
                        HeaderContainsWords                 = $currentBatch
                        SenderAddressLocation               = "HeaderOrEnvelope"
                        ApplyHtmlDisclaimerText             = $ruleHtml
                        ApplyHtmlDisclaimerLocation         = "Prepend"
                        ApplyHtmlDisclaimerFallbackAction   = "Wrap"
                        ExceptIfHeaderContainsMessageHeader  = $ExceptionHeader
                        ExceptIfHeaderContainsWords         = $ExceptionHeaderText
                    }
                    New-ExoRequest -tenantid $TenantFilter -cmdlet "New-TransportRule" -cmdParams $params
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Created rule $ruleName ($($currentBatch.Count) names)" -Sev Info
                }
            }
            catch {
                $ex = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed on rule $ruleName: $($ex.GeneralErrorMessage)" -Sev Error -LogData $ex
            }
        }

        # Exception rule handling
        try {
            $exceptionRule = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-TransportRule" -cmdParams @{ Identity = $exceptionRuleName }
            New-ExoRequest -tenantid $TenantFilter -cmdlet "Set-TransportRule" -cmdParams @{
                Identity  = $exceptionRule.Identity
                Priority  = 0
                Comments  = $exceptionComments
            }
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Updated exception rule (priority 0)." -Sev Info
        }
        catch {
            New-ExoRequest -tenantid $TenantFilter -cmdlet "New-TransportRule" -cmdParams @{
                Name           = $exceptionRuleName
                Comments       = $exceptionComments
                SetHeaderName  = $ExceptionHeader
                SetHeaderValue = $ExceptionHeaderText
                Priority       = 0
                Enabled        = $false
            }
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Created new exception rule (disabled – configure conditions manually before enabling)." -Sev Info
        }

        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Display Name Spoofing protection setup completed successfully." -Sev Info
        return "Display Name Spoofing rules updated successfully for $($displayNames.Count) users across $($batches.Count) rule(s)."
    }
    catch {
        $ex = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Critical failure: $($ex.GeneralErrorMessage)" -Sev Error -LogData $ex
        return "Failed: $($ex.GeneralErrorMessage)"
    }
}