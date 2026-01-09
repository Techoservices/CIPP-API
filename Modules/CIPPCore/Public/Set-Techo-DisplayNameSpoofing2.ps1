<#
.SYNOPSIS
    Updates Exchange Transport Rules to warn on external emails matching internal Display Names.

.DESCRIPTION
    Designed to run via CIPP Script Scheduler or as a Remediation action.
    1. Checks/Updates the Exception rule (High Priority).
       - CREATES as DISABLED if missing (Safety).
       - PRESERVES state if already exists.
    2. Fetches all User AND Shared Mailboxes.
    3. Flushes old "Display Name" rules.
    4. Recreates rules in batches (shards) of 200 names.

    Parameters provided via CIPP:
      - TenantFilter: Tenant ID/domain.
      - Headers: CIPP API Headers.
#>

function Set-Techo-DisplayNameSpoofing {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantFilter,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $false)]
        [string]$APIName = 'Set Display Name Spoofing Rules'
    )

    try {
        $startTime = Get-Date
        $results = [PSCustomObject]@{
            UsersFound = 0
            RulesRemoved = 0
            RulesCreated = 0
            ExceptionRuleUpdated = $false
        }

        # Constants
        $RulePrefix = "Add Disclaimer for Internal Display Names"
        $ExceptionRuleName = "Exceptions to Display Name Disclaimer"
        $ExceptionHeader = "X-Positive-IgnoreDisplayName"
        $ExceptionHeaderValue = "True"
        
        # HTML Disclaimer - CLOSING TAG IS FLUSH LEFT TO PREVENT PARSING ERRORS
        $RuleHtml = @"
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width="100%" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'> <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#ff6200;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width="60%" style='width:98.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding="10px 88px 25px 88px" color="#000000"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:5.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><strong><b><i><span style='font-size:15.0pt;font-family: "Segoe UI",sans-serif;mso-fareast-font-family:"Times New Roman";color:#fc0303'>This message was sent from outside the company by someone with a display name matching a user in your organization. Please do not click any links, or open attachments, unless you recognize the source of this email and know the content is safe. <o:p></o:p></span><strong><b><i></p></div></td></tr></table>
"@

        # ----------------------------------------------------------------
        # SECTION 0: State Check (Fetch ALL Rules Once)
        # ----------------------------------------------------------------
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Fetching current Transport Rules..." -Sev Info
        try {
            $AllRules = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-TransportRule' -ErrorAction Stop
            if (-not $AllRules) { $AllRules = @() } 
        }
        catch {
             $errorMessage = Get-CippException -Exception $_
             Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed to fetch rules: $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
             throw "Failed to fetch current rules"
        }

        # ----------------------------------------------------------------
        # SECTION 1: Manage Exception Rule
        # ----------------------------------------------------------------
        # Base params for both New and Set
        $ExceptParams = @{
            Name            = $ExceptionRuleName
            Comments        = "Matches emails that bypass the Display Name spoof check."
            SetHeaderName   = $ExceptionHeader
            SetHeaderValue  = $ExceptionHeaderValue
        }

        $ExistingExceptionRule = $AllRules | Where-Object { $_.Name -eq $ExceptionRuleName }

        try {
            if ($ExistingExceptionRule) {
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Exception rule found. Updating configuration (State preserved)..." -Sev Debug
                
                # Clone params and add Identity for Set
                # We DO NOT add 'Enabled' here, so existing state (True/False) is preserved
                $SetParams = $ExceptParams.Clone()
                $SetParams.Add("Identity", $ExceptionRuleName)
                
                New-ExoRequest -tenantid $TenantFilter -cmdlet 'Set-TransportRule' -cmdParams $SetParams | Out-Null
                $results.ExceptionRuleUpdated = "Updated"
            } else {
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Exception rule not found. Creating as DISABLED..." -Sev Debug
                
                # Clone params and add Enabled=$false for safety on creation
                $NewParams = $ExceptParams.Clone()
                $NewParams.Add("Enabled", $false)
                
                New-ExoRequest -tenantid $TenantFilter -cmdlet 'New-TransportRule' -cmdParams $NewParams | Out-Null
                $results.ExceptionRuleUpdated = "Created (Disabled)"
            }
        }
        catch {
             $err = Get-CippException -Exception $_
             Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed to Manage Exception Rule: $($err.GeneralErrorMessage)" -Sev Error
        }

        # ----------------------------------------------------------------
        # SECTION 2: Fetch User and Shared Mailboxes
        # ----------------------------------------------------------------
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Fetching User and Shared Mailboxes..." -Sev Info
        
        try {
            $DisplayNames = New-ExoRequest -tenantid $TenantFilter -cmdlet 'Get-Mailbox' -cmdParams @{ 
                ResultSize = 'unlimited'
                RecipientTypeDetails = @('UserMailbox', 'SharedMailbox')
            } -ErrorAction Stop | Select-Object -ExpandProperty DisplayName
        }
        catch {
             $errorMessage = Get-CippException -Exception $_
             Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed to retrieve mailboxes: $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
             throw "Failed to retrieve mailboxes"
        }

        if (-not $DisplayNames) {
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "No mailboxes found. Aborting." -Sev Info
            return "No mailboxes found."
        }
        
        $results.UsersFound = $DisplayNames.Count
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Found $($results.UsersFound) names (Users + Shared)." -Sev Info


        # ----------------------------------------------------------------
        # SECTION 3: Flush Old Spoofing Rules
        # ----------------------------------------------------------------
        $RulesToRemove = $AllRules | Where-Object { $_.Name -like "$RulePrefix*" }
        
        if ($RulesToRemove) {
            Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Removing $($RulesToRemove.Count) old rules..." -Sev Info
            foreach ($Rule in $RulesToRemove) {
                try {
                    New-ExoRequest -tenantid $TenantFilter -cmdlet 'Remove-TransportRule' -cmdParams @{
                        Identity = $Rule.Guid
                        Confirm = $false
                    } -ErrorAction Stop | Out-Null
                    $results.RulesRemoved++
                }
                catch {
                    Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed to remove old rule $($Rule.Name). Continuing." -Sev Warning
                }
            }
        }

        # ----------------------------------------------------------------
        # SECTION 4: Create Batched Rules
        # ----------------------------------------------------------------
        $BatchSize = 200
        $Counter = 1
        $TotalNames = $results.UsersFound

        for ($i = 0; $i -lt $TotalNames; $i += $BatchSize) {
             $EndIndex = [Math]::Min($i + $BatchSize - 1, $TotalNames - 1)
             [string[]]$Batch = $DisplayNames[$i..$EndIndex]
             
             $CurrentRuleName = "$RulePrefix $Counter"

             Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Creating Rule: $CurrentRuleName" -Sev Debug

             $RuleParams = @{
                Name                              = $CurrentRuleName
                FromScope                         = "NotInOrganization"
                HeaderContainsMessageHeader       = "From"
                HeaderContainsWords               = $Batch
                SenderAddressLocation             = "HeaderOrEnvelope"
                ApplyHtmlDisclaimerText           = $RuleHtml
                ApplyHtmlDisclaimerLocation       = "Prepend"
                ApplyHtmlDisclaimerFallbackAction = "Wrap"
                ExceptIfHeaderContainsMessageHeader = $ExceptionHeader
                ExceptIfHeaderContainsWords       = $ExceptionHeaderValue
            }

            try {
                New-ExoRequest -tenantid $TenantFilter -cmdlet 'New-TransportRule' -cmdParams $RuleParams -ErrorAction Stop | Out-Null
                $results.RulesCreated++
            }
            catch {
                $err = Get-CippException -Exception $_
                Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Failed to create rule $CurrentRuleName - $($err.GeneralErrorMessage)" -Sev Error
            }
            $Counter++
        }

        # Generate summary
        $totalTime = (Get-Date) - $startTime
        $formattedTime = "{0:mm}m {0:ss}s" -f $totalTime
        $summary = "Completed in $formattedTime. Names Protected: $($results.UsersFound). Rules Created: $($results.RulesCreated). Exceptions: $($results.ExceptionRuleUpdated)."
        
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message $summary -Sev Info
        
        return $summary
    }
    catch {
        $errorMessage = Get-CippException -Exception $_
        Write-LogMessage -headers $Headers -API $APIName -tenant $TenantFilter -message "Critical Failure: $($errorMessage.GeneralErrorMessage)" -Sev Error -LogData $errorMessage
        throw
    }
}