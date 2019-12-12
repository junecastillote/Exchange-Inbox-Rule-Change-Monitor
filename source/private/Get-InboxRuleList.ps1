Function Get-InboxRuleList {
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$OldRules,
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$CurrentRules
    )

    $FinalResult = @()
    $DeletedRule = @()
    $NewRule = @()
    $NormalRule = @()

    # Find removed rules
    # If Rule is present in Old and not present in Current, it is considered removed
    foreach ($Rule in $OldRules) {
        if (!($CurrentRules | Where-Object { $_.Identity -EQ $Rule.Identity })) {
            $DeletedRule += $Rule
        }
        # else {
        #     $NormalRule += $Rule
        # }
    }

    # Find added rules
    # If Rule is present in Current and not present in Old, it is considered Added
    foreach ($Rule in $CurrentRules) {
        if (!($OldRules | Where-Object { $_.Identity -EQ $Rule.Identity })) {
            $NewRule += $Rule
        }
        else {
            $NormalRule += $Rule
        }
    }

    $NewRule | Add-Member -MemberType NoteProperty -Value "New" -Name RuleState -Force
    $DeletedRule | Add-Member -MemberType NoteProperty -Value "Removed" -Name RuleState -Force
    $NormalRule | Add-Member -MemberType NoteProperty -Value "Normal" -Name RuleState -Force

    $FinalResult = ($NewRule + $DeletedRule + $NormalRule)

    return $FinalResult
}