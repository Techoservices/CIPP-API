<#
.SYNOPSIS
    Optimized script to apply/update transport rules that add a warning disclaimer for potential display name spoofing attacks in Exchange Online.

.DESCRIPTION
    This script is optimized for safe, idempotent execution:
    - Assumes an active Exchange Online session (e.g., when run via CIPP's "Run a command" feature or similar app-only context).
    - Retrieves current mailbox display names, deduplicates, and sorts for consistency.
    - Batches names into groups of 200 (safe limit for HeaderContainsWords).
    - Updates existing rules where possible, creates new rules if more batches are needed, removes extra rules if fewer.
    - Sets the exception rule to highest priority (0) and disclaimer rules to lower priorities (1+).
    - Includes comprehensive try/catch error handling with informative output.
    - Suitable for periodic execution to keep rules in sync with user additions/removals.

    To run periodically: Use CIPP's executive commands manually, or add as a custom standard/scheduled task if you extend CIPP.

.NOTES
    Author: Optimized based on original CyberDrain script
    Requirements: ExchangeOnlineManagement module, existing EXO session (app-only authentication).
#>

# Static disclaimer HTML (kept from original, with warning styling)
$ruleHtml = @"
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width="100%" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#ff6200;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width="60%" style='width:98.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding="10px 88px 25px 88px" color="#000000"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:5.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><strong><b><i><span style='font-size:15.0pt;font-family: "Segoe UI",sans-serif;mso-fareast-font-family:"Times New Roman";color:#fc0303'>This message was sent from outside the company by someone with a display name matching a user in your organization. Please do not click any links, or open attachments, unless you recognize the source of this email and know the content is safe. <o:p></o:p></span></strong></b></i></p></div></td></tr></table>
"@

$ruleNamePrefix = "Add Disclaimer for Internal Display Names"
$ExceptionHeader = "X-Positive-IgnoreDisplayName"
$ExceptionHeaderText = "True"
$exceptionRuleName = "Exceptions to Display Name Disclaimer"
$exceptionComments = "This rule is for matching emails that will not be checked for Display Name spoofing. Please use appropriate caution when adding rules, and try to make sure excepted emails are appropriately authenticated"
$batchSize = 200

try {
    # Retrieve, clean, deduplicate, and sort display names for deterministic batching
    $displayNames = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
                    Select-Object -ExpandProperty DisplayName |
                    Where-Object { $_ -and $_.Trim() -ne '' } |
                    Sort-Object -Unique

    if ($displayNames.Count -eq 0) {
        Write-Warning "No valid display names found in mailboxes. Exiting."
        return
    }

    Write-Host "Retrieved $($displayNames.Count) unique display names."

    # Create batches
    $batches = @()
    for ($i = 0; $i -lt $displayNames.Count; $i += $batchSize) {
        $batches += ,@($displayNames[$i..($i + $batchSize - 1)])
    }

    Write-Host "Created $($batches.Count) batch(es) of up to $batchSize names each."

    # Get existing disclaimer rules, sorted by trailing number in name
    $existingRules = Get-TransportRule -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name -like "$ruleNamePrefix *" } |
                     Sort-Object { 
                         if ($_.Name -match '\s(\d+)$') { [int]$matches[1] } else { 9999 }
                     }

    $maxIndex = [Math]::Max($batches.Count, $existingRules.Count) - 1

    for ($i = 0; $i -le $maxIndex; $i++) {
        $ruleNumber = $i + 1
        $ruleName = "$ruleNamePrefix $ruleNumber"
        $currentBatch = if ($i -lt $batches.Count) { $batches[$i] } else { $null }

        try {
            if ($i -ge $batches.Count -and $i -lt $existingRules.Count) {
                # Remove extra rules
                $rule = $existingRules[$i]
                Remove-TransportRule -Identity $rule.Identity -Confirm:$false -ErrorAction Stop
                Write-Host "Removed excess rule: $($rule.Name)"
            }
            elseif ($i -lt $existingRules.Count) {
                # Update existing rule
                $rule = $existingRules[$i]
                Set-TransportRule -Identity $rule.Identity `
                                  -Name $ruleName `
                                  -HeaderContainsWords $currentBatch `
                                  -Priority ($i + 1) `
                                  -ErrorAction Stop
                Write-Host "Updated rule $ruleName ($($currentBatch.Count) names)"
            }
            else {
                # Create new rule
                $params = @{
                    Name                               = $ruleName
                    Priority                           = ($i + 1)
                    FromScope                          = "NotInOrganization"
                    HeaderContainsMessageHeader        = "From"
                    HeaderContainsWords                = $currentBatch
                    SenderAddressLocation              = "HeaderOrEnvelope"
                    ApplyHtmlDisclaimerText            = $ruleHtml
                    ApplyHtmlDisclaimerLocation        = "Prepend"
                    ApplyHtmlDisclaimerFallbackAction  = "Wrap"
                    ExceptIfHeaderContainsMessageHeader = $ExceptionHeader
                    ExceptIfHeaderContainsWords        = $ExceptionHeaderText
                }
                New-TransportRule @params -ErrorAction Stop
                Write-Host "Created new rule $ruleName ($($currentBatch.Count) names)"
            }
        }
        catch {
            Write-Error "Failed to process rule $ruleName: $_"
        }
    }

    # Handle exception rule (highest priority, placeholder for manual configuration)
    try {
        $exceptionRule = Get-TransportRule $exceptionRuleName -ErrorAction Stop
        Set-TransportRule -Identity $exceptionRule.Identity `
                          -Priority 0 `
                          -Comments $exceptionComments `
                          -ErrorAction Stop
        Write-Host "Updated existing exception rule (priority 0)."
    }
    catch {
        New-TransportRule -Name $exceptionRuleName `
                          -Comments $exceptionComments `
                          -SetHeaderName $ExceptionHeader `
                          -SetHeaderValue $ExceptionHeaderText `
                          -Priority 0 `
                          -Enabled $false
        Write-Host "Created new exception rule (disabled - configure conditions manually before enabling)."
    }

    # Final verification
    $coveredNames = Get-TransportRule "$ruleNamePrefix *" -ErrorAction SilentlyContinue |
                    Select-Object -ExpandProperty HeaderContainsWords |
                    Sort-Object -Unique

    if ((Compare-Object $coveredNames $displayNames).Count -eq 0) {
        Write-Host "Success: All display names are covered by the active rules."
    }
    else {
        Write-Warning "Mismatch detected: Some display names may not be covered."
    }
}
catch {
    Write-Error "Critical failure: $_"
}