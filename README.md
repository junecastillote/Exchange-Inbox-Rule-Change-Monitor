# Exchange Inbox Rule Change Monitor

## Getting the baseline rules

To monitor changes, a baseline rule must be generated first. This will server as the reference set of rules.

You can use the Get-InboxRule cmdlet:

`Get-InboxRule -Mailbox <mailboxID> | Export-CliXML C:\BaseLineRules\<mailboxID>.xml`

See the example below:



## Checking for changes in rules


