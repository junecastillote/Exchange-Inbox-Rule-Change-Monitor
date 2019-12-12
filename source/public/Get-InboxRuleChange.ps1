Function Get-InboxRuleChangeReport {
    [CmdletBinding()]
    param(
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $Mailbox,

        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $ReferenceRule
    )

    try {
        Write-Verbose "Getting current inbox rules for mailbox - $Mailbox"
        $DifferenceRule = @(Get-InboxRule -Mailbox $Mailbox -ErrorAction STOP)
        Write-Verbose "Comparing rules. Only rules that are present on both the reference and current rules list will be compared."
        $finalResult = @()
        foreach ($i in $DifferenceRule) {

            $ref = ($ReferenceRule | Where-Object { $_.Identity -eq ($i.Identity) })
            if ($ref) {
                Write-Verbose "Compare rule: $($ref.Name)"
                $finalResult += (Compare-InboxRules -Mailbox $Mailbox -ReferenceRule $ref -DifferenceRule $i)
            }
        }
        return $finalResult
    }
    catch {
        Write-Output $_.Exception.Message
        break;
    }
}