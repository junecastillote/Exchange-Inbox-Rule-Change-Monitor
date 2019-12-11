Function Compare-InboxRules {
    [cmdletbinding()]
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject
    )

    $objprops = (($ReferenceObject | Get-Member -MemberType Property, NoteProperty).Name)
    $objprops += (($DifferenceObject | Get-Member -MemberType Property, NoteProperty).Name)
    $objprops = $objprops | Sort-Object | Select-Object -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        if ($objprop -ne 'RuleState') {
            $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
            if ($diff) {
                $diffprops = [ordered]@{
                    RuleID       = $ReferenceObject.Identity
                    RuleName     = $ReferenceObject.Name
                    PropertyName = $objprop
                    OldValue     = ($diff | Where-Object { $_.SideIndicator -eq '<=' }).$objprop
                    NewValue     = ($diff | Where-Object { $_.SideIndicator -eq '=>' }).$objprop
                }
                $diffs += New-Object PSObject -Property $diffprops
            }
        }
    }
    #if ($diffs) {return ($diffs | Select-Object *)}
    return ($diffs | Select-Object *)
}

#Compare-InboxRule $a $b[1]