Function Compare-InboxRules {
    [cmdletbinding()]
    Param(
        [string]$Mailbox,
        [PSObject]$ReferenceRule,
        [PSObject]$DifferenceRule
    )

    $properties = (($ReferenceRule | Get-Member -MemberType Property, NoteProperty).Name)
    $properties += (($DifferenceRule | Get-Member -MemberType Property, NoteProperty).Name)
    $properties = $properties | Sort-Object | Select-Object -Unique

    $excludeProperty = @('RuleState', 'PSComputerName', 'RunSpaceID', 'PSShowComputerName')

    $diffs = @()
    foreach ($property in $properties) {
        if ($excludeProperty -notcontains $property) {
            $diff = Compare-Object $ReferenceRule $DifferenceRule -Property $property
            if ($diff) {
                $diffprops = [ordered]@{
                    Mailbox      = $Mailbox
                    RuleID       = $ReferenceRule.Identity
                    RuleName     = $DifferenceRule.Name
                    PropertyName = $property
                    OldValue     = ($diff | Where-Object { $_.SideIndicator -eq '<=' }).$property
                    NewValue     = ($diff | Where-Object { $_.SideIndicator -eq '=>' }).$property
                }
                $diffs += New-Object PSObject -Property $diffprops
            }
        }
    }
    return $diffs
}