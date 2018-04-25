
## enables mailbox audit logging for Ben Smith's mailbox
Set-Mailbox -Identity "Ben Smith" -AuditEnabled $true

## enables mailbox audit logging for all user mailboxes in your organization
Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox"} | Set-Mailbox -AuditEnabled $true

## disables mailbox audit logging for Ben Smith's mailbox
Set-Mailbox -Identity "Ben Smith" -AuditEnabled $false

## specifies that the MessageBind and FolderBind actions performed by administrators will be logged for Ben Smith's mailbox
Set-Mailbox -Identity "Ben Smith" -AuditAdmin MessageBind,FolderBind -AuditEnabled $true

## specifies that the SendAs or SendOnBehalf actions performed by delegate users will be logged for Ben Smith's mailbox
Set-Mailbox -Identity "Ben Smith" -AuditDelegate SendAs,SendOnBehalf -AuditEnabled $true

## specifies that the HardDelete action performed by the mailbox owner will be logged for Ben Smith's mailbox
Set-Mailbox -Identity "Ben Smith" -AuditOwner HardDelete -AuditEnabled $true

## Set all options to enable in some mailbox
Set-Mailbox -Identity Admin -AuditOwner Create,SoftDelete,HardDelete,Update,Move,MoveToDeletedItems

## retrieves the auditing settings for all user mailboxes in your organization
Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox"} | Format-List Name,Audit*

## you can run the following commands to display all the audited actions for a specific user logon type
Get-Mailbox <identity of mailbox> | Select-Object -ExpandProperty AuditAdmin
Get-Mailbox <identity of mailbox> | Select-Object -ExpandProperty AuditDelegate
Get-Mailbox <identity of mailbox> | Select-Object -ExpandProperty AuditOwner