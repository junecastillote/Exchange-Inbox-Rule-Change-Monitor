Function Send-InboxRuleChangeReport {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)]
        [psobject]$ReportObject,

        [parameter(Mandatory)]
        [string]$SmtpServer

    )
}