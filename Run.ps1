Remove-Module ExchangeInboxRuleChangeMonitor -ErrorAction SilentlyContinue
Import-Module $PSScriptRoot\ExchangeInboxRuleChangeMonitor.psd1

$ReferenceRule = @(Import-Clixml .\june.xml)

Get-InboxRuleChangeReport -Mailbox june@poshlab.ml -ReferenceRule $ReferenceRule