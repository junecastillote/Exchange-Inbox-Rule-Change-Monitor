Remove-Module ExchangeInboxRuleChangeMonitor -ErrorAction SilentlyContinue
#Import-Module $PSScriptRoot\ExchangeInboxRuleChangeMonitor.psd1
Import-Module .\ExchangeInboxRuleChangeMonitor.psd1

$Mailbox = 'june@poshlab.ml'
$ReferenceRule = @(Import-Clixml .\june.xml)
$ReportObject = Get-InboxRuleChange -Mailbox $Mailbox -ReferenceRule $ReferenceRule
$ReportFile = ".\reports\report.html"

Write-InboxRuleChangeHTMLReport -ReportObject $ReportObject -Path $ReportFile

$mailProps = @{
    SmtpServer = 'smtp.office365.com'
    Port = 587
    Credential = $credential
    Subject = "Inbox Rule Change Detected on Mailbox - $($Mailbox)"
    Body = (Get-Content $ReportFile -raw)
    From = 'june@poshlab.ml'
    To = 'june@poshlab.ml'
    UseSSL = $true
    BodyAsHtml = $true
}

Send-MailMessage @mailProps