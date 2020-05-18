# Create Office 365 Security Group
# - https://docs.microsoft.com/en-us/powershell/module/msonline/new-msolgroup?view=azureadps-1.0
New-MsolGroup -DisplayName "Tenant_Admins" -Description "Admin Group for Office 365"

# Add admin accounts to group
# - https://www.adamfowlerit.com/2015/08/adding-multiple-cloud-users-to-an-azureoffice-365-security-group/
$group = get-msolgroup -All | where {$_.Displayname -eq “Tenant_Admins”}
$users = get-msoluser -All | select userprincipalname,objectid | where {$_.userprincipalname -like “*admin*”}
$users | foreach {add-msolgroupmember -groupobjectid $group.objectid -groupmembertype “user” -GroupMemberObjectId $_.objectid}
get-msolgroupmember -groupobjectid $group.objectid

# Add security group to have access to all mailboxes
# - https://www.quadrotech-it.com/blog/grant-full-access-to-all-mailboxes-in-office-365/
Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox')} | 
Add-MailboxPermission -User Tenant_Admins -AccessRights FullAccess -InheritanceType all

